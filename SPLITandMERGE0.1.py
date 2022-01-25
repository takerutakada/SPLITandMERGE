import os
import glob
import openpyxl as px
import PyPDF2
import pikepdf
import time
import traceback
from tqdm import tqdm

'''Confidential information'''
from CInfo import DVLPR
''''''

###共通関数
def load(file):
    while True:
        try:
            wb = px.load_workbook(file)
            break
        except PermissionError:
            input(os.path.basename(file) + 'を閉じてからEnterキーを押してください：')
    return wb.worksheets[0]
##PDFパスワード解除関数
def unlock_pwd(file):
    pwd = input(os.path.basename(file) + 'は保護されています。パスワードを入力してください：')
    while True:
        try:
            pdf = pikepdf.open(file, password=pwd)
            break
        except pikepdf._qpdf.PasswordError:
            pwd = input('パスワードが違います。入力し直してください：')
    unlock = pikepdf.new()
    unlock.pages.extend(pdf.pages)
    unlock.save(os.path.splitext(os.path.basename(file))[0] + '_copied.pdf')
    return PyPDF2.PdfFileReader(os.path.splitext(os.path.basename(file))[0] + '_copied.pdf')

###メイン処理
try:
    print('「manual.txt」をご確認いただいた上で実行してください')
    mode = int(input('行う処理を選択してください【分割→0／結合→1】：'))
    ##分割
    if mode == 0:
        print('PDF分割処理を行います')
        input('「SPLIT」フォルダ内に\n・分割したいPDF　※1ファイルのみ\n・分割後のファイル名リスト（filenames.xlsx）\nを保管した後、Enterキーを押してください：')
        while True:
            sp_dic = 'SPLIT'
            sp_pdf = glob.glob(os.path.join(sp_dic, '*.pdf'))
            sp_xlsx = os.path.join(sp_dic, 'filenames.xlsx')
            if len(sp_pdf) == 0:
                input('PDFが保管されていません。やり直してEnterキーを押してください：')
            elif len(sp_pdf) > 1:
                input('PDFが複数保管されています。1つだけ保管してからEnterキーを押してください：')
            elif not os.path.isfile(sp_xlsx):
                input('「filenames.xlsx」が保管されていません。やり直してEnterキーを押してください：')
            else:
                sp_pdf = sp_pdf[0]
                break
        print(os.path.basename(sp_pdf) + 'に対し処理を行います')
        files = [ cell.value for cell in load(sp_xlsx)['A'] ]
        try:
            reader = PyPDF2.PdfFileReader(sp_pdf)
            page_num = reader.getNumPages()
        except PyPDF2.utils.PdfReadError:
            reader = unlock_pwd(sp_pdf)
            page_num = reader.getNumPages()
            sp_pdf = os.path.splitext(os.path.basename(sp_pdf))[0] + '_copied.pdf'
        msg = 'PDFのページ数：' + str(page_num) + '\nfilenames.xlsx内のファイル数：' + str(len(files)) + '\n'
        sp_num = int(page_num / len(files))
        if sp_num < 1:
            print(msg + 'ページ数が足りません。処理を終了します')
        elif page_num % len(files) != 0:
            print(msg + '割り切れません。処理を終了します')
        else:
            save_dic = os.path.join(sp_dic, os.path.splitext(os.path.basename(sp_pdf))[0] + '_splited')
            os.mkdir(save_dic)
            i = 0
            for page in tqdm(range(0, page_num, sp_num)):
                merger = PyPDF2.PdfFileMerger()
                start = page
                end = start + sp_num
                merger.append(sp_pdf, pages=(start,end), import_bookmarks=False)
                file_name = files[i]
                s = os.path.join(save_dic, file_name)
                merger.write(s)
                merger.close()
                i += 1        
            print('分割処理が完了しました')
        if os.path.exists(os.path.basename(sp_pdf)):
            os.remove(os.path.basename(sp_pdf))
    ##結合
    elif mode == 1:
        print('PDF結合処理を行います')
        input('「MERGE>pdf_files」フォルダ内に\n・結合したいPDF　※2ファイル以上\n・結合順序指定ファイル（order.xlsx）　※入力がない場合はデフォルトで指定\nを保管した後、Enterキーを押してください：')
        while True:
            mg_dic = os.path.join('MERGE','pdf_files')
            mg_pdf = glob.glob(os.path.join(mg_dic, '*.pdf'))
            mg_xlsx = os.path.join(mg_dic, 'order.xlsx')
            if len(mg_pdf) == 0 or len(mg_pdf) == 1:
                input('保管されているPDFが0または1つだけです。やり直してEnterキーを押してください：')
            elif not os.path.isfile(mg_xlsx):
                input('「order.xlsx」が保管されていません。やり直してEnterキーを押してください：')
            else:
                break
        try:
            files = [ os.path.join(mg_dic, cell.value) for cell in load(mg_xlsx)['A'] ]
        except TypeError:
            print(os.path.basename(mg_xlsx) + 'に入力がありません。以下の順序で結合します')
            print([os.path.basename(f) for f in mg_pdf])
            files = mg_pdf
        file_name = input('結合後のファイル名を入力してください（拡張子不要）：')
        writer = PyPDF2.PdfFileWriter()
        for mg_file in tqdm(files):
            reader = PyPDF2.PdfFileReader(mg_file)
            try:
                for i in range(reader.getNumPages()):
                    writer.addPage(reader.getPage(i))
            except PyPDF2.utils.PdfReadError:
                reader = unlock_pwd(mg_file)
                for i in range(reader.getNumPages()):
                    writer.addPage(reader.getPage(i))
                os.remove(os.path.splitext(os.path.basename(mg_file))[0] + '_copied.pdf')              
        with open(os.path.join('MERGE', file_name + ".pdf"), "wb") as f:
            writer.write(f)
        print('結合処理が完了しました')
    else:
        print('入力が不正です')
except Exception:
    print(f'エラー！下記エラーコードを{DVLPR}までお知らせください')
    print(traceback.format_exc())
    input('終了するにはEnterキーを押してください')
finally:
    print('ソフトを終了します')
    time.sleep(3)
