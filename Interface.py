import tkinter.filedialog
from tkinter import *
from TargetMaquina import *
from TargetUsuarios import *
from AD import *
from pathlib import Path
import webbrowser


class Window:
    def __init__(self, janela):
        canvas = Canvas(
            janela,
            bg="#3c3c3b",
            height=480,
            width=720,
            bd=0,
            highlightthickness=0,
            relief="ridge")
        canvas.place(x=0, y=0)

        self.background_img = PhotoImage(file=f"img/background.png")
        self.background = canvas.create_image(
            328.0, 173.0,
            image=self.background_img)

        self.img0 = PhotoImage(file=f"img/img0.png")
        self.b0 = Button(
            image=self.img0,
            borderwidth=0,
            highlightthickness=0,
            command=self.gerar_target,
            relief="flat")

        self.b0.place(
            x=425, y=391,
            width=162,
            height=46)

        self.img1 = PhotoImage(file=f"img/img1.png")
        self.b1 = Button(
            image=self.img1,
            borderwidth=0,
            highlightthickness=0,
            command=self.set_csv_usuarios,
            relief="flat")

        self.b1.place(
            x=367, y=328,
            width=92,
            height=27)

        self.img2 = PhotoImage(file=f"img/img2.png")
        self.b2 = Button(
            image=self.img2,
            borderwidth=0,
            highlightthickness=0,
            command=self.set_csv_maquinas,
            relief="flat")

        self.b2.place(
            x=367, y=292,
            width=92,
            height=27)

        self.img3 = PhotoImage(file=f"img/img3.png")
        self.b3 = Button(
            image=self.img3,
            borderwidth=0,
            highlightthickness=0,
            command=self.abrir_sop,
            relief="flat")

        self.b3.place(
            x=677, y=10,
            width=27,
            height=27)

        self.entry0_img = PhotoImage(file=f"img/img_textBox0.png")
        self.entry0_bg = canvas.create_image(
            185.0, 373.0,
            image=self.entry0_img)

        self.entry_dias_limite = Entry(
            bd=0,
            bg="#ffffff",
            highlightthickness=0)

        self.entry_dias_limite.place(
            x=116.0, y=363,
            width=138.0,
            height=18)

        self.entry1_img = PhotoImage(file=f"img/img_textBox1.png")
        self.entry1_bg = canvas.create_image(
            185.0, 296.0,
            image=self.entry1_img)

        self.entry_nome_cliente = Entry(
            bd=0,
            bg="#ffffff",
            highlightthickness=0)

        self.entry_nome_cliente.place(
            x=116.0, y=286,
            width=138.0,
            height=18)

        self.entry2_img = PhotoImage(file=f"img/img_textBox2.png")
        self.entry2_bg = canvas.create_image(
            569.5, 306.0,
            image=self.entry2_img)

        self.entry_csv_maquina = Entry(
            bd=0,
            bg="#ffffff",
            highlightthickness=0)

        self.entry_csv_maquina.place(
            x=462, y=296,
            width=215,
            height=18)

        self.entry3_img = PhotoImage(file=f"img/img_textBox3.png")
        self.entry3_bg = canvas.create_image(
            569.5, 342.0,
            image=self.entry3_img)

        self.entry_csv_usuario = Entry(
            bd=0,
            bg="#ffffff",
            highlightthickness=0)

        self.entry_csv_usuario.place(
            x=462, y=332,
            width=215,
            height=18)

    def abrir_sop(self):
        webbrowser.open('https://softwareone.sharepoint.com/teams/BRSLMManaged/Shared%20Documents/Forms/AllItems.aspx?'
                        'CT=1651500624712&FolderCTID=0x012000964A68C63F090B49938C61AE9A88DB1E&id=%2Fteams%'
                        '2FBRSLMManaged%2FShared%20Documents%2FGeneral%2FTarget%20Generator%2FSOP%2FSOP%20%2D%20Target'
                        '%20generator%5Fv2%2Epdf&parent=%2Fteams%2FBRSLMManaged%2FShared%20Documents%2FGeneral%2FTarget'
                        '%20Generator%2FSOP')

    def get_nome_cliente(self):
        nome_cliente = self.entry_nome_cliente.get()
        return nome_cliente

    def get_qnty_dias(self):
        qnty_dias = self.entry_dias_limite.get()
        return qnty_dias

    def set_csv_maquinas(self):
        self.entry_csv_maquina.delete(0, END)
        csv_maquina = ""
        csv_maquinas = tkinter.filedialog.askopenfiles()
        for data in csv_maquinas:
            csv_maquina += data.name + "?"
        self.entry_csv_maquina.insert(0, str(csv_maquina))

    def get_csv_maquinas(self):
        csv_maquinas = self.entry_csv_maquina.get()
        return csv_maquinas

    def set_csv_usuarios(self):
        self.entry_csv_usuario.delete(0, END)
        csv_usuario = ""
        csv_usuarios = tkinter.filedialog.askopenfiles()
        for data in csv_usuarios:
            csv_usuario += data.name + "?"
        self.entry_csv_usuario.insert(0, str(csv_usuario))

    def get_csv_usuarios(self):
        csv_usuarios = self.entry_csv_usuario.get()
        return csv_usuarios

    def gerar_diretorio_targets(self, diretorio, nome_cliente):
        caminho = str(diretorio)
        cliente = str(nome_cliente)
        pasta = Path(r''+caminho + '/Target-'+cliente)
        pasta.mkdir()
        new_diretorio = caminho + '/Target-' + cliente
        return new_diretorio

    def gerar_target_maquina(self, nome_cliente, diretorio, qnty_dias):
        csv_maquina = self.get_csv_maquinas()
        criar_estruturacao_target_maquina(nome_cliente, diretorio, qnty_dias)
        dados_maquina = ler_csv_maquina(csv_maquina)
        dados_maquina = converter_dados(dados_maquina)
        inserir_dados_maquina(dados_maquina, qnty_dias, diretorio, nome_cliente)
        finalizar_planilha_maquina(diretorio, nome_cliente, qnty_dias)

    def gerar_target_usuario(self, nome_cliente, diretorio, qnty_dias):
        csv_usuario = self.get_csv_usuarios()
        criar_estruturacao_target_usuario(nome_cliente,  diretorio, qnty_dias)
        dados_usuario = ler_csv_usuario(csv_usuario)
        dados_usuario = converter_dados(dados_usuario)
        inserir_dados_usuario(dados_usuario, qnty_dias, diretorio, nome_cliente)
        finalizar_planilha_usuario(diretorio, nome_cliente, qnty_dias)

    def gerar_target(self):
        try:
            nome_cliente = self.get_nome_cliente()
            qnty_dias = self.get_qnty_dias()
            diretorio = tkinter.filedialog.askdirectory()
            if (self.entry_csv_usuario.get() == '') | (self.entry_csv_usuario.get() == ' '):
                self.gerar_target_maquina(nome_cliente, diretorio, qnty_dias)
            elif(self.entry_csv_maquina.get() == '') | (self.entry_csv_maquina.get() == ' '):
                self.gerar_target_usuario(nome_cliente, diretorio, qnty_dias)
            else:
                diretorio = self.gerar_diretorio_targets(diretorio, nome_cliente)

                # Máquinas

                self.gerar_target_maquina(nome_cliente, diretorio, qnty_dias)

                # Usuários

                self.gerar_target_usuario(nome_cliente, diretorio, qnty_dias)

            tkinter.messagebox.showinfo('Target', 'Target successfully generated!')
            ask = tkinter.messagebox.askquestion('Target', 'Exit application?', icon='info')
            if ask == 'yes':
                gui.destroy()

        except:
            tkinter.messagebox.showinfo('Target', message='ERROR')
            raise KeyError


gui = Tk()
window = Window(gui)
gui.title('Target Generator')
gui.geometry("720x480")
gui.config(background='#3c3c3b')
gui.iconbitmap('logo.ico')
gui.mainloop()

