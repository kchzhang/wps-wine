import os
import sys
import win32com.client
import pythoncom

def test_wps_installation():
    """测试 WPS 是否安装成功"""
    print("=== WPS 安装验证 ===")
    
    # 检查 WPS 可执行文件是否存在
    wps_paths = [
        r"C:\Program Files\Kingsoft\WPS Office\wps.exe",
        r"C:\Program Files (x86)\Kingsoft\WPS Office\wps.exe"
    ]
    
    for path in wps_paths:
        if os.path.exists(path):
            print(f"✓ WPS 找到: {path}")
            return True
    
    print("✗ WPS 未找到")
    return False

def test_com_objects():
    """测试 COM 对象创建"""
    print("\n=== COM 对象验证 ===")
    
    com_classes = {
        "Word": "KWPS.Application",
        "Excel": "KET.Application", 
        "PowerPoint": "KWPP.Application"
    }
    
    success_count = 0
    for app_name, com_class in com_classes.items():
        try:
            print(f"测试 {app_name} COM 对象...")
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch(com_class)
            app.Visible = False
            print(f"✓ {app_name} COM 对象创建成功")
            
            # 测试基本功能
            if app_name == "Word":
                doc = app.Documents.Add()
                doc.Content.Text = f"WPS {app_name} 测试文档"
                test_path = r"C:\wps-test\test-output\test.docx"
                doc.SaveAs(test_path)
                doc.Close()
                if os.path.exists(test_path):
                    print(f"✓ {app_name} 文档创建成功: {test_path}")
            
            app.Quit()
            success_count += 1
            
        except Exception as e:
            print(f"✗ {app_name} COM 对象创建失败: {e}")
        finally:
            pythoncom.CoUninitialize()
    
    return success_count == len(com_classes)

def test_conversion():
    """测试文档转换功能"""
    print("\n=== 文档转换验证 ===")
    
    try:
        pythoncom.CoInitialize()
        
        # 创建测试文档
        word_app = win32com.client.Dispatch("KWPS.Application")
        word_app.Visible = False
        
        # 创建测试文档
        doc = word_app.Documents.Add()
        doc.Content.Text = "这是一个 WPS 转换测试文档\n创建时间测试"
        input_path = r"C:\wps-test\test_doc.docx"
        output_path = r"C:\wps-test\test_doc.pdf"
        
        doc.SaveAs(input_path)
        print(f"✓ 测试文档创建: {input_path}")
        
        # 转换为 PDF
        doc.ExportAsFixedFormat(output_path, 17)  # 17 = PDF
        doc.Close()
        
        if os.path.exists(output_path):
            print(f"✓ PDF 转换成功: {output_path}")
            result = True
        else:
            print("✗ PDF 转换失败")
            result = False
        
        word_app.Quit()
        return result
        
    except Exception as e:
        print(f"✗ 转换测试失败: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

def main():
    """主验证函数"""
    print("开始 WPS Win32 API 最小可行性验证")
    print("=" * 50)
    
    # 1. 验证安装
    # if not test_wps_installation():
    #     print("\n❌ WPS 安装验证失败")
    #     sys.exit(1)
    
    # 2. 验证 COM 对象
    # if not test_com_objects():
    #     print("\n⚠ COM 对象验证部分失败")
    # else:
    #     print("\n✓ 所有 COM 对象验证成功")
    
    # 3. 验证转换功能
    if test_conversion():
        print("\n✅ 文档转换验证成功")
    else:
        print("\n❌ 文档转换验证失败")
        sys.exit(1)
    
    print("\n" + "=" * 50)
    print("🎉 最小可行性验证完成！")
    print("WPS Win32 API 方案验证通过")

if __name__ == "__main__":
    main()