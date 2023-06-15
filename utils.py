import win32com.client as win32
import win32gui
import os
import re
import glob


hwp = win32.gencache.EnsureDispatch('HWPFRame.HwpObject')
# hwnd = win32gui.FindWindow(None, 'Noname 1 - HWP')

word = win32.gencache.EnsureDispatch('Word.application')
word.Visible = False


def convert_hwp_to_pdf(path_hwp, path_save_pdf, hwp=hwp):
    hwp.Open(path_hwp)
    hwp.SaveAs(path_save_pdf, 'PDF')

def convert_word_to_pdf(path_word, path_save_pdf, word=word):
    doc = word.Documents.Open(path_word)
    doc.SaveAs(path_save_pdf, FileFormat=17)
    doc.Close()


def get_abs_path(path_source):
    __base_name = os.path.basename(path_source)
    __file_name, ext = os.path.splitext(__base_name)
    pdf_path = os.path.join('out', f'{__file_name}.pdf')

    path_source = os.path.abspath(path_source)
    pdf_path = os.path.abspath(pdf_path)    

    return ext, path_source, pdf_path

def get_files(data_dir='data'):
    files = glob.glob(os.path.join(data_dir, '*'))
    return files 

def create_pdf(_path_file):
    _ext, _path_source, _pdf_path = get_abs_path(_path_file)

    print(f'before:: {_path_source}')
    print(f'after:: {_pdf_path}')

    if 'hwp' in _ext:
        convert_hwp_to_pdf(_path_source, _pdf_path)

    elif 'doc' in _ext:
        convert_word_to_pdf(_path_source, _pdf_path)

if __name__ == '__main__':
    files = get_files()
    for _, _path_file in enumerate(files):
        try:
            create_pdf(_path_file)

        except Exception as err:
            print(err)

    hwp.Quit()
    word.Quit()

