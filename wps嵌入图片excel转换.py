import zipfile
import re
import os
import shutil
import openpyxl
from openpyxl.drawing.image import Image
import xml.etree.ElementTree as ET

def wps_image_converter(input_xlsx, output_xlsx):
    # 创建一个临时文件夹用来做图片中转缓存
    temp_dir = "temp_wps_images"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    print(f"正在加载表格: {input_xlsx} ... (可能需要几秒钟)")
    
    # 打开表格（注意：必须开启 data_only=False，以读取公式）
    try:
        wb = openpyxl.load_workbook(input_xlsx, data_only=False)
        sheet = wb.active
    except Exception as e:
        print(f"打开 Excel 失败，请检查文件是否被占用: {e}")
        return

    # 1. 查找所有包含 DISPIMG 的单元格
    cell_to_id = {}
    for row in sheet.iter_rows():
        for cell in row:
            cell_val = str(cell.value)
            if "DISPIMG" in cell_val:
                match = re.search(r'DISPIMG\s*\(\s*"([^"]+)"', cell_val)
                if match:
                    cell_to_id[cell.coordinate] = match.group(1)

    if not cell_to_id:
        print("表格中未检测到含有 WPS 嵌入图片的单元格。")
        return

    print(f"成功定位到 {len(cell_to_id)} 个包含图片的单元格，正在深入底层提取图源...")

    # 2. 从压缩包读取 WPS 特有的图片映射 XML
    id_to_target = {}
    try:
        with zipfile.ZipFile(input_xlsx, 'r') as z:
            cellimages_bytes = z.read('xl/cellimages.xml')
            rels_bytes = z.read('xl/_rels/cellimages.xml.rels')

            # 解析 .rels 提取 rId -> 物理路径
            rId_to_target = {}
            rels_root = ET.fromstring(rels_bytes)
            for elem in rels_root.iter():
                if elem.tag.endswith('Relationship'):
                    rid = elem.attrib.get('Id')
                    target = elem.attrib.get('Target')
                    if rid and target:
                        rId_to_target[rid] = target

            # 解析 cellimages.xml 提取图片 ID -> rId
            cellimages_root = ET.fromstring(cellimages_bytes)
            for child in cellimages_root:
                name = None
                embed = None
                for sub in child.iter():
                    for k, v in sub.attrib.items():
                        if k.endswith('name') and str(v).startswith('ID_'):
                            name = v
                        if k.endswith('embed'):
                            embed = v
                if name and embed and embed in rId_to_target:
                    id_to_target[name] = rId_to_target[embed]

            print("\n开始自动清空假公式并植入真图片...")
            success_count = 0
            
            # 3. 提取图片并直接插回 openpyxl 的工作表中
            for coord, img_id in cell_to_id.items():
                target_path = id_to_target.get(img_id)
                if not target_path:
                    continue

                # 拼装安全的 zip 内绝对路径
                if target_path.startswith('/'):
                    full_target_path = target_path[1:]
                elif target_path.startswith('xl/'):
                    full_target_path = target_path
                else:
                    full_target_path = f"xl/{target_path}"

                # 提取图片到临时文件夹
                ext = full_target_path.split('.')[-1]
                temp_img_path = os.path.join(temp_dir, f"{coord}.{ext}")
                
                try:
                    with z.open(full_target_path) as source, open(temp_img_path, 'wb') as target_file:
                        shutil.copyfileobj(source, target_file)

                    # 将物理图片转换为 Excel 支持的格式并插入
                    img = Image(temp_img_path)
                    
                    # 【图片大小调整区】 (默认把宽和高固定为90像素，避免过大遮挡)
                    img.width = 90
                    img.height = 90
                    
                    sheet.add_image(img, coord)

                    # 最关键的一步：把这个单元格里的 =_xlfn.DISPIMG 公式彻底删掉
                    sheet[coord].value = ""

                    success_count += 1
                    print(f"--> [OK] 单元格 {coord} 图片已成功修复")

                except Exception as e:
                    print(f"--> [失败] 处理单元格 {coord} 时报错: {e}")
                    
    except KeyError:
        print("解析文件底层失败，未能找到 WPS 的底层图片容器。")
        return

    # 4. 全部处理完后，另存为大家都能打开的新表格
    print(f"\n正在保存转化后的新文件: {output_xlsx} ...")
    wb.save(output_xlsx)
    
    # 5. 可选：用完即毁，清除刚才当做跳板的提取图片文件夹 (如果你想看原图，可以把下行注释掉)
    shutil.rmtree(temp_dir, ignore_errors=True)
    
    print(f"==================================================")
    print(f"完美收工！共修复 {success_count} 张图片。")
    print(f"原始文件完好无损，修复后的文件请查看：{output_xlsx}")

# ================= 运行区 =================
# 原来的文件
input_file = 'aaa.xlsx'
# 想要生成的新文件
output_file = 'bbb.xlsx'

wps_image_converter(input_file, output_file)
