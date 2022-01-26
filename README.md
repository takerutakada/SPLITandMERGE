# SPLITandMERGE

**◆ソフト概要**  
当ソフトは指定したPDFの分割・結合を行うソフトです。  
○分割  
指定したPDFを、別途自身で作成する「namelist.xlsx」に入力されているファイル名に基づいて分割します。  

○結合  
指定したPDF群を、別途自身で作成する「order.xlsx」に入力されているファイル順に結合します。  
***
**◆ファイル構成**
当ソフトは以下のファイル構成の時挙動します。
＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
SPLITandMERGE/ ※  
　├ MERGE/←結合時に使用  
　│　└ pdf_files/  
　│　　├ 【この場所に結合したいPDF群を保管】  
　│　　└ order.xlsx←結合順序を指定するファイル     
　├ SPLIT/←分割時に使用  
　│　├【この場所に分割したいPDFを保管】  
　│　└ filename.xlsx←分割後のファイル名を指定するファイル  
　├ manual.txt←当マニュアル  
　└ SPLITandMERGE.exe←メインソフト  
＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊  
※「SPLITandMERGE」フォルダ自体はどこにおいても構いません。  
　またフォルダ名も変更可能です。  
***
**◆使用方法**
「SPLITandMERGE.exe」をダブルクリックすると黒い画面（コンソール）が現れます。  
コンソールに「行う処理を選択してください【分割→0／結合→1】」と表示されるので、  
分割処理を行いたいときは0、結合処理を行いたいときは1（いずれも半角）を入力してEnterキーを押します。  

○分割  
1.「SPLIT」直下に、分割したいPDFを置きます。  
2.同フォルダにある「namelist.xlsx」のA列に分割後のファイル名を入力しておきます。  
　　例）4ページあるPDFを、「分割後001.pdf」「分割後002.pdf」という各2ページの2ファイルに分割したい。  
　　→「namelist.xlsx」のA1セルに「分割後001.pdf」、A2セルに「分割後002.pdf」と入力します。  
3.コンソールの指示に従い、処理を実行します。  
4.「SPLIT」直下に、「（指定したPDF名）_splited」というフォルダが生成されます。  
　　当フォルダ内に指定した分割ファイルが保管されていることを確認します。  

○結合  
1.「MERGE」>「pdf_files」直下に、結合したいPDF群を置きます。  
2.同フォルダにある「order.xlsx」にファイル名を結合したい順に入力しておきます。  
3.結合後のファイル名をコンソールに入力します。  
4.コンソールの指示に従い、処理を実行します。  
5.「MERGE」直下に、指定した順番に結合されたPDFが３．で入力したファイル名で保管されていることを確認します。  
***
◆注意  
○共通  
処理するPDFが保護されている場合は都度パスワードの入力が求められます。  
パスワードが不明の場合は処理できません。  

○分割  
以下2つのパターンの場合は分割処理は行われません。  
*分割するPDFのページ数が「namelist.xlsx」に入力したファイル数より少ない場合  
*分割するPDFのページ数を「namelist.xlsx」に入力したファイル数で割ると余りが生じる場合  

○結合  
「order.xlsx」に入力がない場合、自動で指定された順序で結合します。  
（結合順序は処理開始前にコンソール上に表示されます）  

途中で処理を中断したい場合、「ctrl + c」でコンソールを閉じることができます。  
