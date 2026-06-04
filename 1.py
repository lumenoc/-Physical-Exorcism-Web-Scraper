import pandas as pd
import time
import os
import tkinter as tk
from tkinter import filedialog
from playwright.sync_api import sync_playwright

def select_file():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    print("📁 正在呼出文件选择窗口，请选择你的 Excel 表格...")
    return filedialog.askopenfilename(
        title="请选择需要处理的 Excel 文件",
        filetypes=[("Excel files", "*.xlsx")]
    )

# 1. 选文件读表格
file_path = select_file()
if not file_path:
    exit()

df = pd.read_excel(file_path, sheet_name=0)
medical_keywords = ['医', '药', '护理', '卫生', '生命', '健康', '口腔', '临床', '针灸']
pattern = '|'.join(medical_keywords)

mask_notnull = df['院系名称'].notnull()
mask_medical = df['院系名称'].astype(str).str.contains(pattern, na=False)

target_indices = df[mask_notnull & mask_medical].index
total_count = len(target_indices)

print(f"\n🚀 目标锁定：{total_count} 个医学类院系。即将唤醒浏览器...")
print("-" * 50)

# 2. 唤醒浏览器自动化搜索
with sync_playwright() as p:
    browser = p.chromium.launch(headless=False) 
    page = browser.new_page()
    
    current_count = 0
    for idx in target_indices:
        current_count += 1
        uni = df.loc[idx, '高校名称']
        dept = df.loc[idx, '院系名称']
        
        query = f"{uni} {dept} 官方网站"
        print(f"[{current_count}/{total_count}] 正在搜索: {uni} - {dept}")
        
        try:
            # 【关键修改 1】：换用国内访问最快的百度
            # 这里的 timeout 设为 30000 (30秒)，给网络足够的时间加载
            page.goto(f"https://www.baidu.com/s?wd={query}", timeout=30000)
            
            # 【关键修改 2】：等待百度的搜索结果主容器加载出来
            # 如果网页白屏，它会一直等到出结果，或者等满10秒才放弃
            page.wait_for_selector('div#content_left', timeout=10000)
            
            found = False
            # 【关键修改 3】：寻找百度特有的标题标签 h3 a
            items = page.query_selector_all('h3.t a')
            
            for item in items:
                title = item.inner_text()
                url = item.get_attribute('href')
                
                # 只要不是空链接，就抓下来
                if url and url.startswith('http'):
                    df.loc[idx, '更正后或者新找的的院系名称'] = f"【参考标题】{title}"
                    df.loc[idx, '院系设置的展示页面链接'] = url
                    print(f" -> ✅ 成功抓取: {title}")
                    found = True
                    break 
            
            if not found:
                print(" -> ⚠️ 网页已加载，但未找到有效链接")
                
        except Exception as e:
            # 如果百度都加载不出来，说明网络极度卡顿或者弹了安全验证
            print(f" -> ❌ 访问受阻。请确认浏览器里是不是弹出了安全验证。")
            time.sleep(5) 
            
        time.sleep(2)
        
    browser.close()

# 3. 保存结果
dir_name = os.path.dirname(file_path)
output_path = os.path.join(dir_name, "医学类院系核查结果_百度直连版.xlsx")
df.to_excel(output_path, index=False)

print("-" * 50)
print(f"🎉 全部完成！结果已保存至: \n{output_path}")