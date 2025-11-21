from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
import aiofiles
import os
import uuid
import time
import asyncio
from typing import List
from .converter import converter
from .models import ConvertRequest, ConvertResponse, ConvertFormat, HealthResponse

app = FastAPI(
    title="WPS文档转换服务",
    description="基于WPS COM组件的文档转换REST API",
    version="1.0.0"
)

# 文件存储配置
UPLOAD_DIR = "/tmp/uploads"
OUTPUT_DIR = "/tmp/outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.on_event("startup")
async def startup_event():
    """服务启动时初始化WPS"""
    print("正在初始化WPS转换服务...")
    if converter.initialize():
        print("WPS转换服务初始化成功")
    else:
        print("WPS转换服务初始化失败")

@app.on_event("shutdown")
async def shutdown_event():
    """服务关闭时清理资源"""
    print("正在关闭WPS转换服务...")
    converter.shutdown()

@app.get("/", response_model=HealthResponse)
async def health_check():
    """健康检查端点"""
    return HealthResponse(
        status="healthy",
        version="1.0.0",
        wps_available=converter.initialized
    )

@app.post("/convert/{format}", response_model=ConvertResponse)
async def convert_document(
    format: ConvertFormat,
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """文档转换接口"""
    
    if not converter.initialized:
        raise HTTPException(status_code=503, detail="WPS服务不可用")
    
    # 生成唯一文件名
    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_DIR, f"{file_id}_{file.filename}")
    output_filename = f"{os.path.splitext(file.filename)[0]}.{format.value}"
    output_path = os.path.join(OUTPUT_DIR, f"{file_id}_{output_filename}")
    
    start_time = time.time()
    
    try:
        # 保存上传文件
        async with aiofiles.open(input_path, 'wb') as out_file:
            content = await file.read()
            await out_file.write(content)
        
        # 执行转换
        success = converter.convert_document(input_path, output_path, format)
        
        conversion_time = time.time() - start_time
        
        if success and os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            
            # 添加后台清理任务
            background_tasks.add_task(cleanup_files, input_path, output_path)
            
            return ConvertResponse(
                success=True,
                message="转换成功",
                output_file=f"/download/{file_id}_{output_filename}",
                file_size=file_size,
                conversion_time=conversion_time
            )
        else:
            raise HTTPException(status_code=500, detail="文档转换失败")
            
    except Exception as e:
        # 清理临时文件
        cleanup_files(input_path, output_path)
        raise HTTPException(status_code=500, detail=f"转换过程出错: {str(e)}")

@app.get("/download/{filename}")
async def download_file(filename: str):
    """下载转换后的文件"""
    file_path = os.path.join(OUTPUT_DIR, filename)
    
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="文件不存在")
    
    return FileResponse(
        path=file_path,
        filename=filename.split('_', 1)[1] if '_' in filename else filename,
        media_type='application/octet-stream'
    )

@app.post("/convert-batch", response_model=List[ConvertResponse])
async def convert_batch_documents(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    format: ConvertFormat = ConvertFormat.PDF
):
    """批量文档转换接口"""
    
    if not converter.initialized:
        raise HTTPException(status_code=503, detail="WPS服务不可用")
    
    tasks = []
    for file in files:
        task = asyncio.create_task(convert_single_file(file, format, background_tasks))
        tasks.append(task)
    
    results = await asyncio.gather(*tasks, return_exceptions=True)
    
    # 处理异常结果
    processed_results = []
    for result in results:
        if isinstance(result, Exception):
            processed_results.append(ConvertResponse(
                success=False,
                message=str(result)
            ))
        else:
            processed_results.append(result)
    
    return processed_results

async def convert_single_file(
    file: UploadFile,
    format: ConvertFormat,
    background_tasks: BackgroundTasks
) -> ConvertResponse:
    """处理单个文件转换"""
    return await convert_document(format, background_tasks, file)

def cleanup_files(*file_paths):
    """清理临时文件"""
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except:
            pass

@app.get("/supported-formats")
async def get_supported_formats():
    """获取支持的转换格式"""
    return {
        "supported_formats": [format.value for format in ConvertFormat],
        "input_formats": [".doc", ".docx", ".txt", ".rtf"],
        "default_format": "pdf"
    }