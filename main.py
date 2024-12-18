import tkinter as tk
import pandas as pd
from time import sleep
from tkinter import scrolledtext, filedialog
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

        self.submit_button = tk.Button(
            root, text="Iniciar o BOT", command=self.main, bg='green', fg='white', width=20)
        self.submit_button.grid(row=2, column=0, padx=10, pady=5, sticky='e')

        self.export_button = tk.Button(
            root, text="Exportar Log", command=self.export_log, bg='blue', fg='white', width=20)
        self.export_button.grid(row=2, column=1, padx=10, pady=5, sticky='w')

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

            # Acessa a transação IW41 e insere a ordem de manutenção
            # session.findById("wnd[0]").resizeWorkingPane(179, 24, False)
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW41"
            session.findById("wnd[0]").sendVKey(0)

            session.findById("wnd[0]/usr/ctxtCORUF-AUFNR").text = row['Ordem']
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[1]/usr/btnOPTION2").press()

            # Verifica se "Tempo de acionamento" está vazio
            if pd.isnull(row['Tempo de acionamento']):
                self.output_text.insert(
                    tk.END, 'Tempo de acionamento está vazio!\n')
                self.output_text.see(tk.END)
                self.root.update_idletasks()
                sleep(1)
            else:
                session.findById(
                    "wnd[0]/usr/tblSAPLCORUTC_3100").getAbsoluteRow(2).selected = True
                self.output_text.insert(tk.END, 'Item 1 selecionado!\n')
                self.output_text.see(tk.END)
                self.root.update_idletasks()
                sleep(1)

            # Verifica se "Taxa de vazamento (m³/min)" está vazio
            if pd.isnull(row['Taxa de vazamento (m³/min)']):
                self.output_text.insert(
                    tk.END, 'Taxa de vazamento (m³/min) está vazio!\n')
                self.output_text.see(tk.END)
                self.root.update_idletasks()
                sleep(1)
            else:
                session.findById(
                    "wnd[0]/usr/tblSAPLCORUTC_3100").getAbsoluteRow(3).selected = True
                self.output_text.insert(tk.END, 'Item 2 selecionado!\n')
                self.output_text.see(tk.END)
                self.root.update_idletasks()
                sleep(1)

            # Clica no botao "Dados reais"
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[1]/usr/btnOPTION2").press()

            # Preenche os campos da transação
            session.findById("wnd[0]/usr/txtAFRUD-ISMNW_2").text = '0'
            session.findById(
                "wnd[0]/usr/ctxtAFRUD-ISDD").text = row["Inicio trabalho"]
            session.findById(
                "wnd[0]/usr/ctxtAFRUD-IEDD").text = row["Fim trabalho"]
            session.findById(
                "wnd[0]/usr/txtAFRUD-LTXA1").text = row["Texto da Confirmacao (Maximo 40 caracteres)"]

            # Clica no botao "Documentos de medicao"
            session.findById("wnd[0]/tbar[1]/btn[9]").press()

            # Verifica se foi exibido uma mensagem de aviso no rodapé e confirma
            error_message = session.findById("wnd[0]/sbar/pane[0]").text
            if error_message == ERROR_MSG:
                session.findById("wnd[0]").sendVKey(0)
                sleep(1)

            if error_message == ERROR_MSG:
                session.findById("wnd[0]").sendVKey(0)
                sleep(1)

            # Preenche os campos da transação
            session.findById(
                "wnd[0]/usr/ctxtRIMR0-DFDAT").text = row["Data medicao"]
            hr_med = row["Hora medicao"].strftime("%H:%M:%S")
            session.findById("wnd[0]/usr/ctxtRIMR0-DFTIM").text = hr_med
            session.findById(
                "wnd[0]/usr/txtRIMR0-DFRDR").text = row["Lido por"]

            # Seleciona o local de instalacao e clica em "Todos ptos.med.obj."
            session.findById("wnd[0]/usr/txtRIMR0-MPOBK").setFocus()
            session.findById("wnd[0]/usr/txtRIMR0-MPOBK").caretPosition = 20
            session.findById("wnd[0]/tbar[1]/btn[25]").press()

            # Valida se o tempo de acionamento está vazio e preenche o campo no local correto
            if pd.isnull(row['Tempo de acionamento']):
                self.output_text.insert(
                    tk.END, 'Tempo de acionamento está vazio!\n')
            else:
                for i in range(0, 8, 2):
                    try:
                        if session.findById(f"wnd[0]/usr/sub:SAPLIMR0:4210/txtIMPT-PTTXT[{i},37]").text == "REGISTRAR TEMPO ACIONAMENTO":
                            session.findById(
                                f"wnd[0]/usr/sub:SAPLIMR0:4210/txtRIMR0-RDCNT[{i+1},3]").text = row["Tempo de acionamento"]
                    except Exception as e:
                        self.output_text.insert(
                            tk.END, f"Linha {i} nao possui o texto padrao.\n")
                        self.output_text.see(tk.END)
                        self.root.update_idletasks()

            # Valida se a taxa de vazamento está vazia e preenche o campo no local correto
            if pd.isnull(row['Taxa de vazamento (m³/min)']):
                self.output_text.insert(
                    tk.END, 'Taxa de vazamento (m³/min) está vazio!\n')
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
            else:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                session.findById("wnd[0]/tbar[1]/btn[23]").press()
                session.findById("wnd[1]/usr/btnOPTION2").press()
                session.findById("wnd[0]/tbar[1]/btn[9]").press()
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/txtRIMR0-MPOBK").setFocus()
                session.findById(
                    "wnd[0]/usr/txtRIMR0-MPOBK").caretPosition = 20
                session.findById("wnd[0]/tbar[1]/btn[25]").press()

                for i in range(0, 8, 2):
                    try:
                        if session.findById(f"wnd[0]/usr/sub:SAPLIMR0:4210/txtIMPT-PTTXT[{i},37]").text == "REGISTRAR TAXA VAZAMENTO GAS":
                            session.findById(f"wnd[0]/usr/sub:SAPLIMR0:4210/txtRIMR0-RDCNT[{
                                             i+1},3]").text = row["Taxa de vazamento (m³/min)"]
                    except Exception as e:
                        self.output_text.insert(
                            tk.END, f"Linha {i} nao possui o texto padrao.\n")
                        self.output_text.see(tk.END)
                        self.root.update_idletasks()

                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
                self.output_text.insert(
                    tk.END, f"{row['Ordem']} - ordem salva com sucesso!\n")
                self.output_text.see(tk.END)
                self.root.update_idletasks()

            self.output_text.see(tk.END)
            self.root.update_idletasks()

        self.output_text.insert(tk.END, "BOT finalizado.\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()

        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


    def export_log(self):
        log_content = self.output_text.get("1.0", tk.END)
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, "w") as file:
                file.write(log_content)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
