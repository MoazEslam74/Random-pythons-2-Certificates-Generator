import os
import win32com.client
from docxtpl import DocxTemplate
from typing import List

lang='EN'
def toggleLang():
    lang='AR' if lang=='EN' else 'EN'
def mainInterface():
    num_certificates=int(input('Enter the number of certificates:'))
    name_list=[]

    for i in range(num_certificates):
        name_list.append(input(f'Enter the name {i+1}:'))
    word_path=input('Enter the path of the word template:')
    tag=input('Enter the tag of changing:')
    output_folder=input('Enter output folder:')

    generate_certificates(names_list=name_list,template_path=word_path,tag_name=tag,output_folder=output_folder)

def generate_certificates(
    names_list: List[str], 
    template_path: str, 
    tag_name: str = "name", 
    output_folder: str = "Certificates"
):
    
    # 1. Check that the template path exists
    if not os.path.exists(template_path):
        print(f"❌ خطأ: قالب الوورد غير موجود في المسار: {template_path}" if lang=='AR' else f'Error the word file doesn\'t exists is path {template_path}')
        return
        
    if not names_list:
        print("⚠️ القائمة فارغة، لم يتم تمرير أي أسماء لإصدار الشهادات." if lang=='AR' else 'Erorr the list is empty, No name has passed to the generator')
        return

    # 2. Create the output folder if it does not exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    # 3. Convert paths to absolute paths; this is required for Word COM
    abs_template_path = os.path.abspath(template_path)
    abs_output_folder = os.path.abspath(output_folder)

    print("🚀 جاري تجهيز برنامج Word في الخلفية..." if lang=='AR' else "🚀 Word working in the background...")
    
    # Use DispatchEx to open an independent session and prevent interference
    try:
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False # forbbide any tab appear
    except Exception as e:
        print(f"❌ حدث خطأ أثناء تشغيل برنامج Word: {e}"if lang=='AR' else "❌ Error happend while Word working")
        return

    try:
        for index, person_name in enumerate(names_list, start=1):
            print(f"⏳ جاري معالجة الشهادة ({index}/{len(names_list)}): {person_name}")
            
            # ----- Templating phase -----
            doc = DocxTemplate(abs_template_path)
            # Pass the tag name dynamically
            context = {tag_name: person_name}
            doc.render(context)
            
            # Prepare temporary and final file paths
            temp_docx_path = os.path.join(abs_output_folder, f"temp_{person_name}.docx")
            pdf_path = os.path.join(abs_output_folder, f"RP_Cert_{person_name}.pdf")
            
            # Save the temporary Word file
            doc.save(temp_docx_path)
            
            # ----- PDF conversion phase -----
            word_doc = None
            try:
                word_doc = word_app.Documents.Open(temp_docx_path)
                # 17 is Word's file format code for PDF
                word_doc.SaveAs(pdf_path, FileFormat=17) 
            except Exception as e:
                print(f"❌ خطأ في تحويل شهادة {person_name}: {e}")
            finally:
                # Close the document even if an error occurs
                if word_doc:
                    word_doc.Close(SaveChanges=False)
            
            # Remove the temporary Word file to keep the folder clean
            if os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
                
        print("\n🎉 تمت العملية بنجاح! جميع الشهادات جاهزة في المجلد المحدد.")
        
    except Exception as e:
        print(f"❌ حدث خطأ غير متوقع أثناء المعالجة: {e}")
    finally:
        # Ensure Word does not remain running in memory
        print("🧹 جاري إغلاق Word وتنظيف الذاكرة...")
        word_app.Quit()

# ==========================================
# Example of how to use the function elsewhere
# ==========================================
if __name__ == "__main__":
    
    mainInterface()