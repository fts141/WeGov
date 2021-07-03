from bs4 import BeautifulSoup

class eGovDocx:
    def __init__(self, docx): # 初期化： インスタンス作成時に自動的に呼ばれる
        self.docx = docx     # インスタンス変数 value を宣言する
        
 
    def writeTable(self, soup):
        

    def print_value(self):    # インスタンス変数 value の値を表示する関数
        print(self.value)     # インスタンス変数 value にアクセスし表示する
 
if __name__ == "__main__":
    a = MyClass("123")        # インスタンス a を作成
    b = MyClass("abc") 