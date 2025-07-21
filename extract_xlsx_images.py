import os
import shutil
import re
import logging
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
import argparse

# log
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class XlsxImageExtractor:
    def __init__(self, xlsx_path):
        self.xlsx_path = xlsx_path
        self.temp_dir = Path("temp_extraction")
        self.output_dir = Path("extracted_images")
        self.question_numbers = {}
        self.namespace = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
    def setup_directories(self):
        """設置必要的目錄"""
        for directory in [self.temp_dir, self.output_dir]:
            if directory.exists():
                shutil.rmtree(directory)
            directory.mkdir(parents=True)

    def extract_xlsx(self):
        """使用 zipfile 解壓縮 xlsx 檔案"""
        try:
            logger.info("解壓縮 Excel 檔案...")
            
            # 建立臨時目錄
            self.setup_directories()
            
            # 使用 zipfile 解壓縮
            with zipfile.ZipFile(self.xlsx_path, 'r') as zip_ref:
                # 檢查 ZIP 檔案完整性
                if zip_ref.testzip() is not None:
                    logger.error("Excel 檔案已損壞")
                    return False
                
                # 顯示 ZIP 檔案內容
                logger.info("Excel 檔案包含以下檔案:")
                for item in zip_ref.namelist():
                    logger.info(f"  {item}")
                    
                # 解壓縮所有檔案
                zip_ref.extractall(self.temp_dir)
                
                # 驗證必要檔案是否存在
                required_files = [
                    "xl/worksheets/sheet1.xml",
                    "xl/media"
                ]
                
                for file_path in required_files:
                    full_path = self.temp_dir / file_path
                    if not full_path.exists():
                        logger.warning(f"找不到必要的檔案: {file_path}")
                
                return True
                
        except zipfile.BadZipFile:
            logger.error("無法解壓縮 Excel 檔案，檔案可能已損壞")
            return False
        except Exception as e:
            logger.error(f"解壓縮過程發生錯誤: {str(e)}")
            return False

    def read_worksheet_xml(self):
        """從 XML 直接讀取工作表內容"""
        logger.info("開始讀取工作表 XML...")
        
        try:
            worksheet_path = self.temp_dir / "xl" / "worksheets" / "sheet1.xml"
            if not worksheet_path.exists():
                logger.error("找不到工作表 XML 檔案")
                return
            
            # 讀取 sharedStrings.xml 獲取字串內容
            shared_strings = self._load_shared_strings()
            
            tree = ET.parse(worksheet_path)
            root = tree.getroot()
            
            ns = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # 讀取所有儲存格
            for row in root.findall(".//w:row", ns):
                row_num = int(row.get("r", "0"))
                cells = row.findall(".//w:c", ns)
                
                # 取得 A、B、C 欄的儲存格
                a_cell = next((c for c in cells if c.get("r", "").startswith(f"A{row_num}")), None)
                b_cell = next((c for c in cells if c.get("r", "").startswith(f"B{row_num}")), None)
                c_cell = next((c for c in cells if c.get("r", "").startswith(f"C{row_num}")), None)
                
                if all([a_cell, b_cell, c_cell]):
                    # 取得儲存格值
                    def get_cell_value(cell):
                        if cell.get("t") == "s":  # 字串類型，需要查詢 sharedStrings
                            v = cell.find(".//w:v", ns)
                            if v is not None and v.text:
                                index = int(v.text.strip())
                                return shared_strings.get(index, f"未知字串 {index}")
                        else:  # 數值類型
                            v = cell.find(".//w:v", ns)
                            return v.text.strip() if v is not None and v.text else None
                    
                    question_no = get_cell_value(a_cell)
                    chapter_no = get_cell_value(b_cell)
                    question_text = get_cell_value(c_cell)
                    
                    if all([question_no, chapter_no, question_text]):
                        # 使用 QuestionNo 和 ChapterNo 組合作為唯一識別碼
                        unique_id = f"{chapter_no}_{question_no}"
                        
                        self.question_numbers[unique_id] = {
                            'question_no': question_no,
                            'chapter_no': chapter_no,
                            'number': f"{chapter_no}.{question_no}",
                            'row': row_num,
                            'col': 3,
                            'full_text': question_text
                        }
                        logger.info(f"找到題號: {chapter_no}.{question_no}")
            
            logger.info(f"總共找到 {len(self.question_numbers)} 個題號")
            
        except ET.ParseError as e:
            logger.error(f"解析 XML 時發生錯誤: {str(e)}")
        except Exception as e:
            logger.error(f"讀取工作表時發生錯誤: {str(e)}")
            logger.error(str(e))

    def _load_shared_strings(self):
        """載入 sharedStrings.xml 檔案中的字串資料"""
        shared_strings = {}
        try:
            shared_strings_path = self.temp_dir / "xl" / "sharedStrings.xml"
            if not shared_strings_path.exists():
                logger.warning("找不到 sharedStrings.xml 檔案")
                return shared_strings
            
            tree = ET.parse(shared_strings_path)
            root = tree.getroot()
            
            ns = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # 讀取所有字串
            for i, si in enumerate(root.findall(".//si", ns)):
                t = si.find(".//t", ns)
                if t is not None:
                    shared_strings[i] = t.text
                else:
                    shared_strings[i] = f"空字串 {i}"
            
            logger.info(f"從 sharedStrings.xml 載入了 {len(shared_strings)} 個字串")
            return shared_strings
            
        except Exception as e:
            logger.error(f"載入 sharedStrings.xml 時發生錯誤: {str(e)}")
            return shared_strings

    def analyze_worksheet_for_images(self):
        """從 sheet1.xml 中分析圖片與題號的關聯"""
        logger.info("從 sheet1.xml 分析圖片與題號關聯...")
        
        image_mappings = {}
        try:
            worksheet_path = self.temp_dir / "xl" / "worksheets" / "sheet1.xml"
            if not worksheet_path.exists():
                logger.error("找不到 sheet1.xml 檔案")
                return image_mappings
            
            # 讀取 sharedStrings.xml 獲取字串內容
            shared_strings = self._load_shared_strings()
            
            # 解析 XML
            tree = ET.parse(worksheet_path)
            root = tree.getroot()
            ns = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # 尋找所有含有 #VALUE! 的儲存格（圖片）
            image_cells = []
            for row in root.findall(".//w:row", ns):
                row_num = int(row.get("r", "0"))
                for cell in row.findall(".//w:c", ns):
                    if cell.get("t") == "e":
                        v_element = cell.find(".//w:v", ns)
                        if v_element is not None and v_element.text == "#VALUE!":
                            vm_value = cell.get("vm", "")
                            if not vm_value:
                                continue
                            cell_ref = cell.get("r", "")
                            image_cells.append({
                                "row_num": row_num,
                                "cell_ref": cell_ref,
                                "vm_value": vm_value
                            })
                            logger.info(f"在第 {row_num} 行找到圖片 (vm={vm_value})")
            
            # 建立圖片與題號的關係
            for img_cell in image_cells:
                row_num = img_cell["row_num"]
                vm_value = img_cell["vm_value"]
                
                # 從該行取得題號和章節資訊
                for row in root.findall(f".//w:row[@r='{row_num}']", ns):
                    # 取得 A、B 欄的儲存格（題號和章節）
                    cells = row.findall(".//w:c", ns)
                    a_cell = next((c for c in cells if c.get("r", "").startswith(f"A{row_num}")), None)
                    b_cell = next((c for c in cells if c.get("r", "").startswith(f"B{row_num}")), None)
                    c_cell = next((c for c in cells if c.get("r", "").startswith(f"C{row_num}")), None)
                    
                    # 取得儲存格值
                    question_no = None
                    chapter_no = None
                    question_text = None
                    
                    # 處理 A 欄（題號）
                    if a_cell is not None:
                        if a_cell.get("t") == "s":  # 字串類型
                            v = a_cell.find(".//w:v", ns)
                            if v is not None and v.text:
                                index = int(v.text.strip())
                                question_no = shared_strings.get(index, "")
                        else:  # 數值類型
                            v = a_cell.find(".//w:v", ns)
                            if v is not None and v.text:
                                question_no = v.text.strip()
                    
                    # 處理 B 欄（章節）
                    if b_cell is not None:
                        if b_cell.get("t") == "s":  # 字串類型
                            v = b_cell.find(".//w:v", ns)
                            if v is not None and v.text:
                                index = int(v.text.strip())
                                chapter_no = shared_strings.get(index, "")
                        else:  # 數值類型
                            v = b_cell.find(".//w:v", ns)
                            if v is not None and v.text:
                                chapter_no = v.text.strip()
                    
                    # 處理 C 欄（問題文字）
                    if c_cell is not None:
                        if c_cell.get("t") == "s":  # 字串類型
                            v = c_cell.find(".//w:v", ns)
                            if v is not None and v.text:
                                index = int(v.text.strip())
                                question_text = shared_strings.get(index, "")
                        else:  # 數值類型
                            v = c_cell.find(".//w:v", ns)
                            if v is not None and v.text:
                                question_text = v.text.strip()
                    
                    # 建立圖片映射，使用行號 + VM 值作為唯一鍵
                    mapping_key = f"{row_num}_{vm_value}"
                    image_mappings[mapping_key] = {
                        "row": row_num,
                        "question_no": question_no,
                        "chapter_no": chapter_no,
                        "question_text": question_text,
                        "vm_value": vm_value
                    }
                    
                    logger.info(f"圖片 (行={row_num}, vm={vm_value}) 對應題號 {chapter_no}.{question_no}")
            
            return image_mappings
            
        except ET.ParseError as e:
            logger.error(f"解析 XML 時發生錯誤: {str(e)}")
            return image_mappings
        except Exception as e:
            logger.error(f"分析 worksheet 時發生錯誤: {str(e)}")
            logger.error(str(e))
            return image_mappings

    def process(self, keep_temp=False):
        """處理整個提取和對應過程"""
        try:
            # 使用 zipfile 解壓縮
            if not self.extract_xlsx():
                return None
            
            # 讀取工作表 XML
            self.read_worksheet_xml()
            
            # 分析 sheet1.xml 獲取圖片與題號關聯
            image_info = self.analyze_worksheet_for_images()
            
            # 從 Excel 中提取所有圖片
            logger.info("從 Excel 提取圖片...")
            extracted_images = []
            media_dir = self.temp_dir / "xl" / "media"
            if media_dir.exists():
                # 統計圖片檔案總數
                media_files = [file for file in media_dir.glob("*") 
                              if file.is_file() and file.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif']]
                
                # 複製圖片並建立對應關係
                result = {}
                
                # 按照行號_VM值的對應關係建立映射
                image_mappings = list(image_info.keys())
                logger.info(f"找到 {len(image_mappings)} 個圖片映射: {', '.join(image_mappings)}")
                
                # 確保 media_files 和 image_mappings 有合適數量的元素
                if len(media_files) < len(image_mappings):
                    logger.warning(f"警告: 圖片文件數量 ({len(media_files)}) 少於映射數量 ({len(image_mappings)})")
                
                # 為每個映射分配圖片
                sorted_mappings = sorted(image_mappings, key=lambda k: (int(k.split('_')[0]), int(k.split('_')[1])))
                
                # 按 VM 值分組收集映射
                vm_to_mappings = {}
                for mapping_key in sorted_mappings:
                    row, vm = mapping_key.split('_')
                    if vm not in vm_to_mappings:
                        vm_to_mappings[vm] = []
                    vm_to_mappings[vm].append(mapping_key)
                
                # 遍歷所有 VM 組
                for vm_value, vm_mappings in sorted(vm_to_mappings.items(), key=lambda x: int(x[0])):
                    # 查找對應的圖片文件
                    vm_index = list(sorted(vm_to_mappings.keys())).index(vm_value)
                    
                    if vm_index < len(media_files):
                        image_file = sorted(media_files, key=lambda x: x.name)[vm_index]
                        base_name, ext = os.path.splitext(image_file.name)
                        
                        # 為該 VM 值下的每個映射複製圖片
                        for i, mapping_key in enumerate(vm_mappings):
                            # 如果同一 VM 有多個映射，添加後綴
                            if len(vm_mappings) > 1:
                                image_name = f"{base_name}_{i+1}{ext}"
                            else:
                                image_name = image_file.name
                                
                            dest_path = self.output_dir / image_name
                            shutil.copy2(image_file, dest_path)
                            extracted_images.append(image_name)
                            
                            # 使用映射鍵來查找相關信息
                            info = image_info[mapping_key]
                            result[image_name] = {
                                'cell_position': f"行 {info['row']}",
                                'question_no': info['question_no'],
                                'chapter_no': info['chapter_no'],
                                'unique_id': f"{info['chapter_no']}_{info['question_no']}",
                                'question': f"{info['chapter_no']}.{info['question_no']}",
                                'question_text': info['question_text'],
                                'vm_value': info['vm_value'],
                                'mapping_key': mapping_key,
                                'original_image': image_file.name
                            }
                            logger.info(f"圖片 {image_name} (行={info['row']}, VM={info['vm_value']}) 對應到題號 {info['chapter_no']}.{info['question_no']}")
                    else:
                        for mapping_key in vm_mappings:
                            logger.warning(f"映射 {mapping_key} 沒有對應的圖片文件")
            
            # 處理結果
            report = {
                'total_questions': len(self.question_numbers),
                'total_images': len(extracted_images),
                'total_mappings': len(image_info),
                'question_details': {
                    unique_id: {
                        'question_no': info['question_no'],
                        'chapter_no': info['chapter_no'],
                        'number': info['number'],
                        'full_text': info['full_text']
                    }
                    for unique_id, info in self.question_numbers.items()
                },
                'image_mappings': result,
                'extracted_images': extracted_images
            }
            
            logger.info(f"處理完成！找到 {len(self.question_numbers)} 個題號和 {len(extracted_images)} 張圖片，共 {len(image_info)} 個映射關係")
            return report
            
        except Exception as e:
            logger.error(f"處理過程中發生錯誤: {str(e)}")
            return None
        finally:
            # 清理臨時檔案
            if not keep_temp:
                try:
                    if self.temp_dir.exists():
                        shutil.rmtree(self.temp_dir)
                        logger.info("已清理臨時檔案")
                except Exception as e:
                    logger.error(f"清理臨時檔案時發生錯誤: {str(e)}")

def main():
    
    parser = argparse.ArgumentParser(description='從 Excel 檔案中提取圖片並建立與題號的對應關係')
    parser.add_argument('xlsx_file', help='Excel 檔案路徑')
    parser.add_argument('--debug', action='store_true', help='啟用除錯模式')
    parser.add_argument('--keep-temp', action='store_true', help='保留臨時檔案')
    parser.add_argument('--verbose', '-v', action='store_true', help='顯示詳細結果')
    args = parser.parse_args()
    
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    if not os.path.exists(args.xlsx_file):
        logger.error(f"找不到檔案: {args.xlsx_file}")
        return
    
    extractor = XlsxImageExtractor(args.xlsx_file)
    result = extractor.process(keep_temp=args.keep_temp)
    
    if result:
        logger.info("\n處理結果摘要:")
        logger.info(f"- 總題數: {result['total_questions']}")
        logger.info(f"- 總圖片數: {result['total_images']}")
        
        logger.info("\n題號列表:")
        for q_id, info in result['question_details'].items():
            logger.info(f"- {info['number']}: {info['full_text'][:50]}...")
        
        logger.info("\n圖片對應關係:")
        for img_name, info in result['image_mappings'].items():
            if info['question']:
                logger.info(f"- {img_name}: {info['question']} ({info['cell_position']})")
                if args.verbose:
                    logger.info(f"  題目: {info['question_text'][:100]}...")
            else:
                logger.info(f"- {img_name}: 未對應")
        
        print("\n圖片與題號對應表:")
        print("------------------------------------------------------------------------------")
        print("| 圖片檔案      | 章節號 | 題號 | 對應方式       | 映射鍵        | 原始圖片    |")
        print("------------------------------------------------------------------------------")
        for img_name, info in result['image_mappings'].items():
            chapter = info['chapter_no'] or 'N/A'
            question = info['question_no'] or 'N/A'
            method = info['cell_position']
            mapping_key = info.get('mapping_key', 'N/A')
            original = info.get('original_image', img_name)
            print(f"| {img_name:14} | {chapter:6} | {question:4} | {method:14} | {mapping_key:12} | {original:10} |")
        print("------------------------------------------------------------------------------")
        
        print(f"\n總結: 找到 {result['total_questions']} 個題號, {result['total_images']} 張圖片, {result['total_mappings']} 個映射關係")
    else:
        logger.error("處理失敗")
        
    # 顯示路臨時檔案徑 (選要)
    if args.keep_temp:
        logger.info(f"\n臨時檔案保留在: {os.path.abspath('temp_extraction')}")

if __name__ == "__main__":
    main() 