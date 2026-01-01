import win32com.client
import os
import time

def create_role_matrix(filename):
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Role Matrix"
        
        # Header
        ws.Cells(1, 1).Value = "Role / Permission"
        for i in range(2, 52):
            ws.Cells(1, i).Value = f"Permission {i-1}"
            
        # Data
        for row in range(2, 20):
            ws.Cells(row, 1).Value = f"Role {row-1}"
            for col in range(2, 52):
                ws.Cells(row, col).Value = "X" if (row + col) % 3 == 0 else ""

        # Formatting
        ws.Range(ws.Cells(1, 1), ws.Cells(20, 51)).Borders.LineStyle = 1
        
        wb.SaveAs(os.path.abspath(filename))
        wb.Close()
        excel.Quit()
    except Exception as e:
        print(f"Error creating {filename}: {e}")
        try: excel.Quit() 
        except: pass

def create_complex_layout(filename):
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Screen Description"
        
        # Title
        ws.Cells(1, 1).Value = "Login Screen Specification"
        ws.Cells(1, 1).Font.Size = 20
        ws.Cells(1, 1).Font.Bold = True
        
        # Table Simulation
        ws.Cells(3, 1).Value = "Field"
        ws.Cells(3, 2).Value = "Type"
        ws.Cells(3, 3).Value = "Description"
        ws.Cells(3, 4).Value = "Validation"
        
        data = [
            ("Username", "Text", "Input for user ID", "Required, Email format"),
            ("Password", "Password", "Input for password", "Required, Min 8 chars"),
            ("Login Btn", "Button", "Submits form", "Enabled only if valid"),
            ("Forgot Pwd", "Link", "Redirects to recovery", "-"),
        ]
        
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                ws.Cells(4+i, 1+j).Value = val
                
        ws.Range(ws.Cells(3, 1), ws.Cells(3+len(data), 4)).Borders.LineStyle = 1
        
        # Add a Shape to simulate image/diagram
        # AddShape(Type, Left, Top, Width, Height)
        # msoShapeRectangle = 1
        shape = ws.Shapes.AddShape(1, 50, 200, 300, 150) 
        shape.TextFrame.Characters().Text = "Mockup Image Placeholder"
        
        wb.SaveAs(os.path.abspath(filename))
        wb.Close()
        excel.Quit()
    except Exception as e:
        print(f"Error creating {filename}: {e}")
        try: excel.Quit() 
        except: pass

def create_hidden_text_reproduction(filename):
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Hidden Text"
        
        # Scenario: Long text, WrapText=False, Column Width small
        ws.Cells(1, 1).Value = "Short"
        ws.Cells(1, 2).Value = "Short"
        
        long_text = "This is a very long text that will definitely be hidden if wrap text is not enabled and the column is too narrow."
        ws.Cells(2, 1).Value = "Label"
        ws.Cells(2, 2).Value = long_text
        
        # Force column width small and WrapText False
        ws.Columns(2).ColumnWidth = 10
        ws.Cells(2, 2).WrapText = False
        
        wb.SaveAs(os.path.abspath(filename))
        wb.Close()
        excel.Quit()
    except Exception as e:
        print(f"Error creating {filename}: {e}")
        try: excel.Quit() 
        except: pass

def create_mixed_layout(filename):
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Mixed Layout"
        
        # Scenario: Column A has "Short" (No Wrap) and "Long" (Wrap)
        ws.Cells(1, 1).Value = "Header (No Wrap)"
        ws.Cells(2, 1).Value = "A very long description that should be wrapped and not force column width expansion."
        ws.Cells(3, 1).Value = "Short Item"
        
        # Set WrapText
        ws.Cells(1, 1).WrapText = False
        ws.Cells(2, 1).WrapText = True
        ws.Cells(3, 1).WrapText = False
        
        # Initial bad state: narrow column
        ws.Columns(1).ColumnWidth = 10
        
        wb.SaveAs(os.path.abspath(filename))
        wb.Close()
        excel.Quit()
    except Exception as e:
        print(f"Error creating {filename}: {e}")
        try: excel.Quit() 
        except: pass

def create_test_excel(filename, content="Test Content", wide=False):
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Cells(1, 1).Value = content
        
        if wide:
            # Fill columns A to Z to make it wide
            for i in range(1, 27):
                ws.Cells(1, i).Value = f"Col {i}"
                
        wb.SaveAs(os.path.abspath(filename))
        wb.Close()
        excel.Quit()
    except Exception as e:
        print(f"Error creating {filename}: {e}")
        # Try to quit if stuck
        try:
            excel.Quit()
        except:
            pass

if __name__ == "__main__":
    if not os.path.exists("test_data"):
        os.makedirs("test_data")
    
    create_test_excel("test_data/normal.xlsx", "Normal File")
    time.sleep(1)
    create_test_excel("test_data/wide.xlsx", "Wide File", wide=True)
    time.sleep(1)
    create_test_excel("test_data/japanese.xlsx", "こんにちは")
    time.sleep(1)
    create_test_excel("test_data/vietnamese.xlsx", "Xin chào")
    time.sleep(1)
    
    create_role_matrix("test_data/role_matrix.xlsx")
    time.sleep(1)
    create_complex_layout("test_data/screen_desc.xlsx")
    time.sleep(1)
    create_hidden_text_reproduction("test_data/hidden_text.xlsx")
    time.sleep(1)
    create_mixed_layout("test_data/mixed_layout.xlsx")
    
    print("Test data created.")
