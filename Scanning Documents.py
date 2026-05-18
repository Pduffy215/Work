from docx import Document
import openpyxl
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os 
import win32com.client as win32

def find_highlighted_text(doc):
    equipment_stars = {}
    seen = set()

    for para in doc.paragraphs:
        underlined_runs = []
        
        for run in para.runs:
            if run.font.underline:
                underlined_runs.append(run.text)
            else:
                if underlined_runs:
                    text = "".join(underlined_runs).strip()
                    
                    if text and text not in seen:
                        seen.add(text)
                        
                        star_count = 0
                        for char in text:
                            if char == "*":
                                star_count += 1
                            else:
                                break
                        
                        clean_text = text.lstrip("*").strip()
                        equipment_stars[clean_text] = star_count             
                    underlined_runs = []
        if underlined_runs:
            text = "".join(underlined_runs).strip()
            
            if text and text not in seen:
                seen.add(text)
                
                star_count = 0
                for char in text:
                    if char == "*":
                        star_count += 1
                    else:
                        break
                
                clean_text = text.lstrip("*").strip()
                equipment_stars[clean_text] = star_count

    return equipment_stars

def write_to_excel(data, filename):
    color_map = {
    1: "99CCFF",  # Blue
    2: "FFFF00",  # Yellow
    3: "FFA500",  # Orange
    4: "FF3300",  # Red
    0: "008000"   # Green
}
    
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    for key, value in data.items():
        search_text = key

        for row in sheet.iter_rows(min_row=2, min_col=2, max_col=3, values_only=False):
            cell_equipment = row[0]
            cell_stars = row[1]
            if cell_equipment.value == search_text:
                i = 0
                stars = ""
                while i < value:
                    stars += "*"
                    i = i + 1
                cell_stars.value = stars
                star_count = str(cell_stars.value or "").count("*")
                cell_stars.fill = openpyxl.styles.PatternFill(start_color=color_map.get(star_count, "FFFFFF"), end_color=color_map.get(star_count, "FFFFFF"), fill_type="solid")
                break

    wb.save(filename)

def wait_for_file(file_path, timeout=10):
    start_time = time.time()
    while True:
        try:
            doc = Document(file_path)
            return doc  
        except Exception:
            if time.time() - start_time > timeout:
                raise TimeoutError(f"File not ready: {file_path}")
            time.sleep(0.5)

def convert_doc_to_docx(file_path):
    if file_path.lower().endswith(".doc") and not file_path.lower().endswith(".docx"):
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(file_path)
        new_file = os.path.splitext(file_path)[0] + ".docx"
        doc.SaveAs(new_file, FileFormat=16)  # 16 = .docx
        doc.Close()
        word.Quit()
        print(f"Converted {file_path} -> {new_file}")
        return new_file
    else:
        return file_path

class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
          if not event.is_directory:
            file_path = event.src_path
            print(f"New file detected: {file_path}")

            if file_path.lower().endswith(".doc") and not file_path.lower().endswith(".docx"):
                try:
                    file_path = convert_doc_to_docx(file_path)  # <-- Make sure convert_doc_to_docx is defined
                except Exception as e:
                    print(f"Failed to convert .doc to .docx: {e}")
                    return

            try:
                doc = wait_for_file(file_path)
            except TimeoutError as e:
                print(e)
                return

            equipment_stars = find_highlighted_text(doc)
            print(equipment_stars)
            write_to_excel(equipment_stars, "temp.xlsx")

def main():
    folder_to_watch = r"C:\Users\PCduffy\IncomingFiles"
    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, folder_to_watch, recursive=False)
    observer.start()

    print(f"Watching folder: {folder_to_watch}")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()