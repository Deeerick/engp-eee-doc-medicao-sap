import tkinter as tk
import pandas as pd

from time import sleep
from tkinter import scrolledtext
from utils.sap_logon import login_sap
from utils.config import ERROR_MSG


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("ENGP/EEE - Documentos de Medição")
        self.root.resizable(False, False)

        self.label_key = tk.Label(root, text="Informe a sua chave:")
        self.label_key.grid(row=0, column=0, padx=10, pady=5, sticky='w')

        self.entry_key = tk.Entry(root)
        self.entry_key.grid(row=0, column=1, padx=10, pady=5)
        self.entry_key.focus_set()

        self.label_password = tk.Label(root, text="Informe a sua senha:")
        self.label_password.grid(row=1, column=0, padx=10, pady=5, sticky='w')

        self.entry_password = tk.Entry(root, show='*')
        self.entry_password.grid(row=1, column=1, padx=10, pady=5)

        self.submit_button = tk.Button(root, text="Iniciar o BOT", command=self.main, bg='green', fg='white', width=35)
        self.submit_button.grid(row=2, column=0, columnspan=2, padx=10, pady=5)

        self.output_text = scrolledtext.ScrolledText(root, width=50, height=8)
        self.output_text.grid(row=3, column=0, columnspan=2, padx=10, pady=5)


    def main(self):
        user = self.entry_key.get()
        password = self.entry_password.get()
        self.output_text.insert(tk.END, "Iniciando o BOT...\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()
        
        # Lê a planilha Excel
        df = pd.read_excel('planilha-padrao.xlsx')
        
        # Abre o SAP
        session = login_sap()
        self.output_text.insert(tk.END, 'SAP aberto com sucesso!\n')
        self.output_text.see(tk.END)
        self.root.update_idletasks()
        sleep(5)
        
        # Itera sobre cada linha do DataFrame e imprime os valores das colunas especificadas
        for index, row in df.iterrows():
            try:
                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
                session.findById("wnd[0]").sendVKey(0)
            
            except:
                continue
            
            finally:
                session.findById("wnd[0]").resizeWorkingPane(123, 38, False)
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW41"
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/ctxtCORUF-AUFNR").text = row['Ordem']
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[1]/usr/btnOPTION2").press()
                
                if pd.isnull(row['Inicio trabalho']):
                    self.output_text.insert(tk.END, 'Inicio trabalho está vazio!\n')
                    self.output_text.see(tk.END)
                    self.root.update_idletasks()
                    sleep(1)
                else:
                    session.findById("wnd[0]/usr/tblSAPLCORUTC_3100").getAbsoluteRow(2).selected = True
                    self.output_text.insert(tk.END, 'Item 1 selecionado!\n')
                    self.output_text.see(tk.END)
                    self.root.update_idletasks()
                    sleep(1)
                
                if pd.isnull(row['Fim trabalho']):
                    self.output_text.insert(tk.END, 'Fim trabalho está vazio!\n')
                    self.output_text.see(tk.END)
                    self.root.update_idletasks()
                    sleep(1)
                else:
                    session.findById("wnd[0]/usr/tblSAPLCORUTC_3100").getAbsoluteRow(3).selected = True
                    self.output_text.insert(tk.END, 'Item 2 selecionado!\n')
                    self.output_text.see(tk.END)
                    self.root.update_idletasks()
                    sleep(1)
                
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[1]/usr/btnOPTION2").press()
                
                session.findById("wnd[0]/usr/txtAFRUD-ISMNW_2").text = '0'
                session.findById("wnd[0]/usr/ctxtAFRUD-ISDD").text = row["Inicio trabalho"]
                session.findById("wnd[0]/usr/ctxtAFRUD-IEDD").text = row["Fim trabalho"]
                session.findById("wnd[0]/usr/txtAFRUD-LTXA1").text = row["Texto da Confirmacao (Maximo 40 caracteres)"]
                
                session.findById("wnd[0]/tbar[1]/btn[9]").press()
                
                error_message = session.findById("wnd[0]/sbar/pane[0]").text
                if error_message == ERROR_MSG:
                    session.findById("wnd[0]").sendVKey(0)
                    sleep(1)
                    
                    if error_message == ERROR_MSG:
                        session.findById("wnd[0]").sendVKey(0)
                        sleep(1)
                        
                session.findById("wnd[0]/usr/ctxtRIMR0-DFDAT").text = row["Data medicao"]
                
                hr_med = row["Hora medicao"].strftime("%H:%M:%S")
                session.findById("wnd[0]/usr/ctxtRIMR0-DFTIM").text = hr_med
                
                session.findById("wnd[0]/usr/txtRIMR0-DFRDR").text = row["Lido por"]
                
                sleep(10)
                sleep(500)
                
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW41"
                session.findById("wnd[0]").sendVKey(0)
                
                self.output_text.see(tk.END)
                self.root.update_idletasks()
        
        self.output_text.insert(tk.END, "BOT finalizado.\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()
        
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
