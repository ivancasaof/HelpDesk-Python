from tkinter import *
from tkinter import ttk
from tkinter import messagebox, filedialog
import time, pyodbc, csv, os, shutil
from PIL import ImageTk, Image
from tkinter import scrolledtext
from ldap3 import Server, Connection, ALL, NTLM, ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES, AUTO_BIND_NO_TLS, SUBTREE
import smtplib, ssl, bcrypt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import subprocess, xlsxwriter

fonte_titulo_principal = ('Segoe UI Bold',18,)
fonte_titulos = ('Segoe UI Bold', 12)
fonte_padrao = ('Segoe UI Semibold', 10)
fonte_botao = ('Segoe UI Semibold', 8)
#/////////////////////////////VARIAVEL GLOBAL/////////////////////////////
imgconvertida = None
notifica = None
compara = None
status_atendimento = None
email_interacao = None
anexo_extensao = None

data = time.strftime('%d/%m/%Y', time.localtime())
hora = time.strftime('%H:%M:%S', time.localtime())
global titulo_todos
titulo_todos = 'HelpDesk GV - v4.0'
global versao
versao = '4.0'


global email
email= str("teste")
global controle_loop
controle_loop = 0

global ativa_filtro
ativa_filtro = 0

global pesquisa_com_filtro_tabela
pesquisa_com_filtro_tabela = None

global pesquisa_com_filtro_filtro
pesquisa_com_filtro_filtro = None

global nome_anexo
nome_anexo = None

global caminho_anexo
caminho_anexo = None


def contador():
    cont = cursor.execute("SELECT * FROM dbo.chamados WHERE id_analista=? AND status NOT LIKE '%Encerrado%' AND status NOT LIKE '%Cancelado%' OR status=? ORDER BY prioridade DESC, id_chamado",(usuariologado,'Aberto',))
    for cont in cont.fetchall():
        global notifica
        notifica = cont[0]

def notificacao():
    from win10toast import ToastNotifier
    toaster = ToastNotifier()
    toaster.show_toast("HelpDeskGV", "\nUm novo chamado foi aberto!", duration=10, icon_path="imagens\\ico.ico", threaded=True)

def principal():
    root2.destroy()
    root = Tk()
    root.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
    root.unbind_class("Button", "<Key-space>")
    root.focus_force()
    root.grab_set()
    def duploclique_tree_principal(event):
        if nivel_acesso == 0:
            visualizar_chamado()
        else:
            atendimento()

    def atualizar_lista_principal_encerrado():
        if controle_loop == 0:
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM dbo.chamados WHERE id_analista=? AND status NOT LIKE '%Encerrado%' AND status NOT LIKE '%Cancelado%' OR status=? ORDER BY prioridade DESC, id_chamado",(usuariologado,'Aberto',))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    if row[13] == None:
                        row[13] = ''
                    if row[15] == None:
                        row[15] = ''
                    tree_principal.insert('', 'end', text=" ",
                                          values=(
                                          row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                          tags=('par',))
                else:
                    if row[13] == None:
                        row[13] = ''
                    if row[15] == None:
                        row[15] = ''
                    tree_principal.insert('', 'end', text=" ",
                                          values=(
                                          row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                          tags=('impar',))
                cont += 1
            contador()
        if ativa_filtro == 1:
            atualizar_lista_com_filtro()

    def atualizar_lista_principal():
        print("entrou na lista principal")
        #label criada apenas para o contator ter uma referencia dentro dessa função
        global lbl_loop
        lbl_loop = Label(frame4, text='', bg="#1d366c")
        lbl_loop.grid(row=0, column=2)

        if nivel_acesso == 0:
            btnatendimento.config(state='disabled')
            image_atendimento = Image.open('imagens\\atendimento_over.png')
            resize_atendimento = image_atendimento.resize((30, 35))
            nova_image_atendimento = ImageTk.PhotoImage(resize_atendimento)
            btnatendimento.photo = nova_image_atendimento
            btnatendimento.config(image=nova_image_atendimento, fg='#7c7c7c')
            btnatendimento.unbind("<Enter>")
            btnatendimento.unbind("<Leave>")

            btnferramentas.config(state='disabled')
            image_ferramentas = Image.open('imagens\\ferramentas_over.png')
            resize_ferramentas = image_ferramentas.resize((35, 35))
            nova_image_ferramentas = ImageTk.PhotoImage(resize_ferramentas)
            btnferramentas.photo = nova_image_ferramentas
            btnferramentas.config(image=nova_image_ferramentas, fg='#7c7c7c')
            btnferramentas.unbind("<Enter>")
            btnferramentas.unbind("<Leave>")

            clique_busca.set('Filtrar por...')
            ent_busca.delete(0, END)
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM dbo.chamados WHERE solicitante = ? OR enviar_chamado =? ORDER BY id_chamado DESC", (usuariologado, usuariologado,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    if row[13] == None:
                        row[13] = ''
                    if row[15] == None:
                        row[15] = ''
                    tree_principal.insert('', 'end', text=" ",
                                          values=(row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                          tags=('par',))
                else:
                    if row[13] == None:
                        row[13] = ''
                    if row[15] == None:
                        row[15] = ''
                    tree_principal.insert('', 'end', text=" ",
                                          values=(row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                          tags=('impar',))
                cont += 1

        else:
            btnatendimento.config(state='normal')
            image_atendimento = Image.open('imagens\\atendimento.png')
            resize_atendimento = image_atendimento.resize((30, 35))
            nova_image_atendimento = ImageTk.PhotoImage(resize_atendimento)
            btnatendimento.photo = nova_image_atendimento
            btnatendimento.config(image=nova_image_atendimento, fg='#ffffff')
            btnatendimento.bind("<Enter>", muda_atendimento)
            btnatendimento.bind("<Leave>", volta_atendimento)

            btnferramentas.config(state='normal')
            image_ferramentas = Image.open('imagens\\ferramentas.png')
            resize_ferramentas = image_ferramentas.resize((35, 35))
            nova_image_ferramentas = ImageTk.PhotoImage(resize_ferramentas)
            btnferramentas.photo = nova_image_ferramentas
            btnferramentas.config(image=nova_image_ferramentas, fg='#ffffff')
            btnferramentas.bind("<Enter>", muda_ferramentas)
            btnferramentas.bind("<Leave>", volta_ferramentas)

            clique_busca.set('Filtrar por...')
            ent_busca.delete(0, END)
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM dbo.chamados WHERE id_analista=? AND status NOT LIKE '%Encerrado%' AND status NOT LIKE '%Cancelado%' OR status=? ORDER BY prioridade DESC, id_chamado",(usuariologado,'Aberto',))
            #cursor.execute("SELECT * FROM dbo.chamados WHERE status NOT LIKE '%Encerrado%' AND status NOT LIKE '%Cancelado%' ORDER BY prioridade DESC, id_chamado")
            # cursor.execute("SELECT * FROM dbo.chamados WHERE status LIKE '%Aberto%' OR status LIKE '%Em andamento%' ORDER BY id_chamado DESC")
            # cursor.execute("SELECT * FROM dbo.chamados ORDER BY id_chamado DESC")
            #r = cursor.execute("SELECT * FROM dbo.chamados WHERE id_chamado=?", (n_chamado,))

            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    if row[13] == None:
                        row[13] = ''
                    if row[15] == None:
                        row[15] = ''
                    tree_principal.insert('', 'end', text=" ",
                                          values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                          tags=('par',))
                else:
                    if row[13] == None:
                        row[13] = ''
                    if row[15] == None:
                        row[15] = ''
                    tree_principal.insert('', 'end', text=" ",
                                          values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                          tags=('impar',))
                cont += 1

            comp = cursor.execute("SELECT * FROM dbo.chamados WHERE id_analista=? AND status NOT LIKE '%Encerrado%' AND status NOT LIKE '%Cancelado%' OR status=? ORDER BY prioridade DESC, id_chamado",(usuariologado,'Aberto',))
            for conti in comp.fetchall():
                global compara
                compara = conti[0]
            global notifica
            if compara < notifica:
                notifica = compara
            elif compara > notifica:
                notificacao()
                notifica = compara

        if usuariologado == 'Administrador':
            btnatendimento.config(state='disabled')
            btnchamado.config(state='disabled')

        if controle_loop == 0:
            loop_atualização()
            print("loop ativado")
        else:
            print("loop desativado")
        if ativa_filtro == 1:
            atualizar_lista_com_filtro()
            print("filtro ativado")
        else:
            print("filtro desativado")

    def atualizar_lista_com_filtro():
        print("entrou na lista com filtro")
        if pesquisa_com_filtro_tabela == 'Status':
            busca_personalizada = "SELECT * FROM dbo.chamados WHERE status LIKE '%" + pesquisa_com_filtro_filtro + "%' ORDER BY id_chamado"
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute(busca_personalizada)
            for row in cursor:
                if row[13] == None:
                    row[13] = ''
                if row[15] == None:
                    row[15] = ''
                tree_principal.insert('', 'end', text=" ",
                                  values=(row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                  tags=('par',))
        elif pesquisa_com_filtro_tabela == "Nº Chamado":
            busca_personalizada = "SELECT * FROM dbo.chamados WHERE id_chamado = " + pesquisa_com_filtro_filtro
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute(busca_personalizada)
            for row in cursor:
                if row[13] == None:
                    row[13] = ''
                if row[15] == None:
                    row[15] = ''
                tree_principal.insert('', 'end', text=" ",
                                      values=(
                                      row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                      tags=('par',))
        elif pesquisa_com_filtro_tabela == "Solicitante":
            busca_personalizada = "SELECT * FROM dbo.chamados WHERE solicitante LIKE '%" + pesquisa_com_filtro_filtro + "%' ORDER BY status"
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute(busca_personalizada)
            for row in cursor:
                if row[13] == None:
                    row[13] = ''
                if row[15] == None:
                    row[15] = ''
                tree_principal.insert('', 'end', text=" ",
                                      values=(
                                      row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                      tags=('par',))
        elif pesquisa_com_filtro_tabela == "Ocorrência":
            busca_personalizada = "SELECT * FROM dbo.chamados WHERE ocorrencia LIKE '%" + pesquisa_com_filtro_filtro + "%' ORDER BY status, id_chamado"
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute(busca_personalizada)
            for row in cursor:
                if row[13] == None:
                    row[13] = ''
                if row[15] == None:
                    row[15] = ''
                tree_principal.insert('', 'end', text=" ",
                                      values=(
                                      row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                      tags=('par',))
        elif pesquisa_com_filtro_tabela == "Analista":
            busca_personalizada = "SELECT * FROM dbo.chamados WHERE id_analista LIKE '%" + pesquisa_com_filtro_filtro + "%' ORDER BY status, id_chamado"
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute(busca_personalizada)
            for row in cursor:
                if row[13] == None:
                    row[13] = ''
                if row[15] == None:
                    row[15] = ''
                tree_principal.insert('', 'end', text=" ",
                                      values=(
                                      row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                      tags=('par',))
        elif pesquisa_com_filtro_tabela == "Data Encerramento":
            busca_personalizada = "SELECT * FROM dbo.chamados WHERE data_encerramento LIKE '%" + pesquisa_com_filtro_filtro + "%' ORDER BY id_chamado DESC"
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute(busca_personalizada)
            for row in cursor:
                if row[13] == None:
                    row[13] = ''
                if row[15] == None:
                    row[15] = ''
                tree_principal.insert('', 'end', text=" ",
                                      values=(
                                      row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                      tags=('par',))

    # /////////////////////////////LOGIN INTERNO/////////////////////////////
    def login_interno():
        root2 = Toplevel()
        root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
        root2.unbind_class("Button", "<Key-space>")
        root2.focus_force()
        root2.grab_set()


        def sair():
            root2.destroy()

        def entrar():
            user = euser.get()
            senha = esenha.get()
            r = cursor.execute("SELECT * FROM dbo.analista WHERE login=?", (user,))
            result = r.fetchone()
            if result != None:
                clique.set("Analista")

            if user == "" or senha == "":
                messagebox.showwarning('Login: Erro', 'Digite o Usuário ou Senha.', parent=root2)
            else:
                if clique.get() == "Usuário":
                    server_name = '192.168.1.19'
                    domain_name = 'gvdobrasil'
                    server = Server(server_name, get_info=ALL)
                    try:
                        Connection(server, user='{}\\{}'.format(domain_name, user), password=senha, authentication=NTLM,
                                   auto_bind=True)
                        global nivel_acesso
                        nivel_acesso = 0
                        global usuariologado
                        usuariologado = user
                        atualizar_lista_principal()
                        root2.destroy()
                    except:
                        messagebox.showwarning('Login: Erro', 'Usuário ou senha inválidos.', parent=root2)
                else:
                    r = cursor.execute("SELECT * FROM dbo.analista WHERE login=?", (user,))
                    result = r.fetchone()
                    if result is None:
                        messagebox.showwarning('Login: Erro', 'Usuário ou Senha inválidos.', parent=root2)
                    else:
                        r = cursor.execute("SELECT * FROM dbo.analista WHERE login=?", (user,))
                        for login in r.fetchall():
                            filtro_user = login[1]
                            filtro_pwd = login[3]
                        if bcrypt.checkpw(senha.encode("utf-8"), filtro_pwd.encode("utf-8")):
                            nivel_acesso = 1
                            usuariologado = login[2]
                            contador()
                            atualizar_lista_principal()
                            root2.destroy()
                        else:
                            messagebox.showwarning('Login: Erro', 'Usuário ou Senha inválidos.', parent=root2)

        def entrar_bind(event):
            entrar()
        frame0 = Frame(root2, bg='#ffffff')
        frame0.grid(row=0, column=0, stick='nsew')
        root2.grid_rowconfigure(0, weight=1)
        root2.grid_columnconfigure(0, weight=1)
        frame1 = Frame(frame0, bg="#1d366c")
        frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
        frame2 = Frame(frame0, bg='#ffffff')
        frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=10)
        frame3 = Frame(frame0, bg='#ffffff')
        frame3.pack(side=TOP, fill=X, expand=False, anchor='center')
        frame4 = Frame(frame0, bg='#ffffff')
        frame4.pack(side=TOP, fill=X, expand=False, anchor='center', pady=10)
        frame5 = Frame(frame0, bg='#1d366c')
        frame5.pack(side=TOP, fill=X, expand=False, anchor='center')

        image_login = Image.open('imagens\\login.png')
        resize_login = image_login.resize((35, 35))
        nova_image_login = ImageTk.PhotoImage(resize_login)

        lbllogin = Label(frame1, image=nova_image_login, text=" Trocar Usuário", compound="left", bg='#1d366c',
                         fg='#FFFFFF', font=fonte_titulos)
        lbllogin.photo = nova_image_login
        lbllogin.grid(row=0, column=1)
        frame1.grid_columnconfigure(0, weight=1)
        frame1.grid_columnconfigure(2, weight=1)

        Label(frame2, text="Modo de Acesso:", bg='#ffffff', fg='#000000', font=fonte_padrao).grid(row=0, column=1,
                                                                                                  sticky="w")
        clique = StringVar()
        clique.set("Usuário")
        drop = OptionMenu(frame2, clique, "Usuário", "Analista")
        drop.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                    highlightthickness=0, relief=RIDGE, width=9, font=fonte_padrao, cursor="hand2")
        drop.grid(row=0, column=2, pady=10)
        frame2.grid_columnconfigure(0, weight=1)
        frame2.grid_columnconfigure(3, weight=1)

        Label(frame3, text="Usuário:", bg='#ffffff', fg='#000000', font=fonte_padrao).grid(row=1, column=1, sticky="w")
        euser = Entry(frame3, width=30, font=fonte_padrao)
        euser.grid(row=1, column=2, sticky="w", padx=5, pady=10)
        euser.focus_force()
        euser.bind("<Return>", entrar_bind)
        Label(frame3, text="Senha:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=2, column=1, sticky="w")
        esenha = Entry(frame3, show="*", width=30, font=fonte_padrao)
        esenha.grid(row=2, column=2, sticky="w", padx=5, pady=10)
        esenha.bind("<Return>", entrar_bind)
        frame3.grid_columnconfigure(0, weight=1)
        frame3.grid_columnconfigure(3, weight=1)

        bt1 = Button(frame4, text='Entrar', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                     activeforeground="#FFFFFF", highlightthickness=0, width=10, relief=RIDGE, command=entrar,
                     font=fonte_padrao, cursor="hand2")
        bt1.grid(row=0, column=1, pady=5, padx=5)
        bt2 = Button(frame4, text='Sair', width=10, relief=RIDGE, command=sair, font=fonte_padrao, cursor="hand2")
        bt2.grid(row=0, column=2, pady=5, padx=5)
        frame4.grid_columnconfigure(0, weight=1)
        frame4.grid_columnconfigure(3, weight=1)

        Label(frame5, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1, sticky="ew")
        frame5.grid_columnconfigure(0, weight=1)
        frame5.grid_columnconfigure(2, weight=1)
        '''root2.update()
        largura = frame0.winfo_width()
        altura = frame0.winfo_height()
        print(largura, altura)'''
        window_width = 330
        window_height = 275
        screen_width = root2.winfo_screenwidth()
        screen_height = root2.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
        #root2.resizable(0, 0)
        root2.iconbitmap('imagens\\ico.ico')
        root2.title(titulo_todos)
    # /////////////////////////////FIM LOGIN INTERNO/////////////////////////////

    # /////////////////////////////ABRIR CHAMADO/////////////////////////////
    def abrirchamado_bind(event):
        abrirchamado()
    def abrirchamado():

        root2 = Toplevel(root)
        root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
        root2.unbind_class("Button", "<Key-space>")
        root2.focus_force()
        root2.grab_set()
        hora = time.strftime('%H:%M:%S', time.localtime())

        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ CÂMERAS \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_cameras_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()
            options = [
                "Instalação",
                "Configuração",
                "Mudança de local",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_cameras_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Sem acesso",
                "Imagem travando",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_cameras_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_cameras_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_cameras_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM CÂMERAS \\\\\\\\\\\\\\\\\\\\\\\\\\\

        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ VPN \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_vpn_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()
                if cliquesub.get() == "Novos acessos":
                    txtdescr.delete('1.0', END)
                    txtdescr.insert(END,'Nome Completo:\n\nSetor:\n\nMatricula:\n\nCargo:\n\nGerente/Supervisor:\n\nPerfil Espelho:\n\nLogins necessários:\n(Exemplos: Domínio (Computador),Totvs,E-mail)\n\nGrupo de E-mail:\n\nComputador a ser configurado:\n\nOutros acessos:')
                    txtdescr.focus_force()


            options = [
                "Instalação",
                "Configuração",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_vpn_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Caindo a conexão",
                "Lentidão",
                "Não conecta",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_vpn_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_vpn_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_vpn_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM VPN \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ ACESSOS \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_acessos_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()
                if cliquesub.get() == "Novos acessos":
                    txtdescr.delete('1.0', END)
                    txtdescr.insert(END,'Nome Completo:\n\nSetor:\n\nMatrícula:\n\nCargo:\n\nGerente/Supervisor:\n\nPerfil Espelho:\n\nLogins necessários:\n(Exemplos: Domínio (Computador),Totvs,E-mail)\n\nGrupo de E-mail:\n\nComputador a ser configurado:\n\nOutros acessos:')
                    txtdescr.focus_force()
                if cliquesub.get() == "SGI":
                    txtdescr.delete('1.0', END)
                    txtdescr.insert(END,'Nome Completo:\n\nSIMEC ou Terceiro:\n\nMatrícula:\n\nUnidade Gerencial: (Ex: Laminação, Aciaria, Qualidade, etc...)\n\nÁrea de Atuação: (Referente a área dentro da unidade gerencial. Ex: Forno Panela da Aciaria).\n\nFunção:\n\nGrupo de Permissão de acesso:\n\nE-mail:')
                    txtdescr.focus_force()
                if cliquesub.get() == "Desligamento do Colaborador":
                    txtdescr.delete('1.0', END)
                    txtdescr.insert(END,'Dados colaborador\nNome Funcionário:\nMatrícula:\nSetor:\nData do Desligamento:\n\nAtividades a serem executadas\nBloqueio de Usuário: Dominio ( ) Totvs Protheus ( ) E-mail ( ) Outros ( )\nRedirecionar e-mail (Informar conta):\nResposta automática de e-mail (informar o texto desejado):\nBackup do usuário? Sim () - Não ()')
                    txtdescr.focus_force()


            options = [
                "Liberação de acesso à pastas",
                "Liberação de acesso à midia externa",
                "Liberação de acesso: Outros",
                "Novos acessos",
                "SGI",
                "Desligamento do Colaborador",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_acessos_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Sem permissão (Leitura ou Gravação)",
                "Espaço insuficiente",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_acessos_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_acessos_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_acessos_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM ACESSOS \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ HARDWARE \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_hardware_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Opções",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_hardware_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Máquina desligando",
                "Máquina reiniciando",
                "Não liga",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_hardware_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_hardware_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_hardware_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM HARDWARE \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ TELEFONIA \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_telefonia_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Troca de Ramal",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_telefonia_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Não recebe ligação",
                "Teclado com falha",
                "Não liga",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_telefonia_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_telefonia_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_telefonia_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM TELEFONIA \\\\\\\\\\\\\\\\\\\\\\\\\\\

        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ RADIO \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_radio_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                elif cliquesub.get() == "Envio de rádio para manutenção externa":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.delete('1.0', END)
                    txtdescr.insert(END,'Modelo:\n\nNº do Rádio:\n\nNº de Série:\n\nÁrea:\n\nResponsável:\n\nDefeito:\n\nCausa:')
                    txtdescr.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Compra",
                "Configuração",
                "Envio de rádio para manutenção externa",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")
        def opt_radio_problemas():
            cliquetipo.set('')
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")
        def opt_radio_duvidas():
            cliquetipo.set('')
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")
        def opt_radio_melhorias():
            cliquetipo.set('')
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")
        def opt_radio_projetos():
            cliquetipo.set('')
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")
        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM RADIO \\\\\\\\\\\\\\\\\\\\\\\\\\\



        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ REDE \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_rede_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Opções",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_rede_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Sem acesso a rede (Dados)",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_rede_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_rede_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_rede_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM REDE \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ INTERNET \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_internet_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Desbloqueio de Sites",
                "Liberação WhatsAppWeb",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_internet_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Sem acesso",
                "Lentidão",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_internet_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_internet_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_internet_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM INTERNET \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ IMPRESSORA \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_impressora_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Instalação",
                "Troca de Toner",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_impressora_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Não está imprimindo",
                "Enroscando papel",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_impressora_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_impressora_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_impressora_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM IMPRESSORA \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ E-MAIL \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_email_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Configuração",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_email_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Não envia e-mail",
                "Não recebe e-mail",
                "Pedindo senha",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_email_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_email_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_email_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM E-MAIL \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ SOFTWARES \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_softwares_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Instalação de Software",
                "Opções",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_softwares_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Configuração",
                "Travamento",
                "Opções",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_softwares_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_softwares_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_softwares_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM SOFTWARES \\\\\\\\\\\\\\\\\\\\\\\\\\\
        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ WINDOWS \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_windows_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Opções",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_windows_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Lentidão",
                "Vírus",
                "Opções",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_windows_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_windows_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_windows_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM WINDOWS \\\\\\\\\\\\\\\\\\\\\\\\\\\

        #\\\\\\\\\\\\\\\\\\\\\\\\\\\ PROTHEUS \\\\\\\\\\\\\\\\\\\\\\\\\\\
        def opt_protheus_solicitacao():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Criação de usuário",
                "Liberação de Módulo",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_protheus_problemas():
            dropsub.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.config(state='disabled')
            def clique(event):
                if cliquesub.get() == "Outros assuntos..":
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                else:
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.insert(0, cliquesub.get())
                    entrytitulo.config(state='disabled')
                    txtdescr.focus_force()

            options = [
                "Sem acesso",
                "Trocar senha",
                "Outros assuntos.."
            ]
            cliquesub = StringVar()
            dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
            dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                                     highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub_problemas.grid(row=1, column=3, sticky="w")

        def opt_protheus_duvidas():
            # dropsub_solicitacao.grid_forget()
            # dropsub_problemas.grid_forget()
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_protheus_melhorias():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        def opt_protheus_projetos():
            entrytitulo.config(state='normal')
            entrytitulo.delete(0, END)
            entrytitulo.focus_force()
            cliquesub = StringVar()
            dropsub = OptionMenu(frame3, cliquesub, "")
            dropsub.grid_forget()
            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
            dropsub.grid(row=1, column=3, sticky="w")

        # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM protheus \\\\\\\\\\\\\\\\\\\\\\\\\\\

        def dropselecaotipo(event):
            if clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Solicitação":
                opt_protheus_solicitacao()
            elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Problemas":
                opt_protheus_problemas()
            elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Dúvidas":
                opt_protheus_duvidas()
            elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Melhorias":
                opt_protheus_melhorias()
            elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Projetos":
                opt_protheus_projetos()
            elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Solicitação":
                opt_windows_solicitacao()
            elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Problemas":
                opt_windows_problemas()
            elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Dúvidas":
                opt_windows_duvidas()
            elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Melhorias":
                opt_windows_melhorias()
            elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Projetos":
                opt_windows_projetos()
            elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Solicitação":
                opt_softwares_solicitacao()
            elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Problemas":
                opt_softwares_problemas()
            elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Dúvidas":
                opt_softwares_duvidas()
            elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Melhorias":
                opt_softwares_melhorias()
            elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Projetos":
                opt_softwares_projetos()
            elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Solicitação":
                opt_email_solicitacao()
            elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Problemas":
                opt_email_problemas()
            elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Dúvidas":
                opt_email_duvidas()
            elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Melhorias":
                opt_email_melhorias()
            elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Projetos":
                opt_email_projetos()
            elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Solicitação":
                opt_impressora_solicitacao()
            elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Problemas":
                opt_impressora_problemas()
            elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Dúvidas":
                opt_impressora_duvidas()
            elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Melhorias":
                opt_impressora_melhorias()
            elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Projetos":
                opt_impressora_projetos()
            elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Solicitação":
                opt_internet_solicitacao()
            elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Problemas":
                opt_internet_problemas()
            elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Dúvidas":
                opt_internet_duvidas()
            elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Melhorias":
                opt_internet_melhorias()
            elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Projetos":
                opt_internet_projetos()
            elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Solicitação":
                opt_radio_solicitacao()
            elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Problemas":
                opt_radio_problemas()
            elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Dúvidas":
                opt_radio_duvidas()
            elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Melhorias":
                opt_radio_melhorias()
            elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Projetos":
                opt_radio_projetos()
            elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Solicitação":
                opt_rede_solicitacao()
            elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Problemas":
                opt_rede_problemas()
            elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Dúvidas":
                opt_rede_duvidas()
            elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Melhorias":
                opt_rede_melhorias()
            elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Projetos":
                opt_rede_projetos()
            elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Solicitação":
                opt_telefonia_solicitacao()
            elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Problemas":
                opt_telefonia_problemas()
            elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Dúvidas":
                opt_telefonia_duvidas()
            elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Melhorias":
                opt_telefonia_melhorias()
            elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Projetos":
                opt_telefonia_projetos()
            elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Solicitação":
                opt_hardware_solicitacao()
            elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Problemas":
                opt_hardware_problemas()
            elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Dúvidas":
                opt_hardware_duvidas()
            elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Melhorias":
                opt_hardware_melhorias()
            elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Projetos":
                opt_hardware_projetos()
            elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Solicitação":
                opt_acessos_solicitacao()
            elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Problemas":
                opt_acessos_problemas()
            elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Dúvidas":
                opt_acessos_duvidas()
            elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Melhorias":
                opt_acessos_melhorias()
            elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Projetos":
                opt_acessos_projetos()
            elif clique_ocorr.get() == "VPN" and cliquetipo.get() == "Solicitação":
                opt_vpn_solicitacao()
            elif clique_ocorr.get() == "VPN" and cliquetipo.get() == "Problemas":
                opt_vpn_problemas()
            elif clique_ocorr.get() == "VPN" and cliquetipo.get() == "Dúvidas":
                opt_vpn_duvidas()
            elif clique_ocorr.get() == "VPN" and cliquetipo.get() == "Melhorias":
                opt_vpn_melhorias()
            elif clique_ocorr.get() == "VPN" and cliquetipo.get() == "Projetos":
                opt_vpn_projetos()
            elif clique_ocorr.get() == "Câmeras" and cliquetipo.get() == "Solicitação":
                opt_cameras_solicitacao()
            elif clique_ocorr.get() == "Câmeras" and cliquetipo.get() == "Problemas":
                opt_cameras_problemas()
            elif clique_ocorr.get() == "Câmeras" and cliquetipo.get() == "Dúvidas":
                opt_cameras_duvidas()
            elif clique_ocorr.get() == "Câmeras" and cliquetipo.get() == "Melhorias":
                opt_cameras_melhorias()
            elif clique_ocorr.get() == "Câmeras" and cliquetipo.get() == "Projetos":
                opt_cameras_projetos()

        def dropselecao_ocorr(event):
            if clique_ocorr.get() == "Protheus":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Windows":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Softwares":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "E-mail":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Impressora":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Internet":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Rádio":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Rede":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Telefonia":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Hardware":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Acessos":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "VPN":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
            elif clique_ocorr.get() == "Câmeras":
                entrytitulo.config(state='normal')
                entrytitulo.delete(0, END)
                entrytitulo.config(state='disabled')
                cliquetipo.set('')
                droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=17)
                cliquesub = StringVar()
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.grid_forget()
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")

        def anexo():
            anexo = filedialog.askopenfilename(initialdir="os.path.expanduser(default_dir)", title="Escolha um Arquivo",
                                               filetypes=([("Todos os arquivos", "*.*")]), parent=root2)
            entryanexo.config(state=NORMAL)
            entryanexo.insert(0, anexo)
            entryanexo.config(state=DISABLED)
            l_nome_anexo = os.path.basename(anexo)
            l_caminho_anexo = anexo
            global nome_anexo
            nome_anexo = l_nome_anexo
            global caminho_anexo
            caminho_anexo = l_caminho_anexo
        def salvar():
            if cliquetipo.get() == "" or entrytitulo.get() == "" or txtdescr.get("1.0", 'end-1c') == "" or clique_setor.get() == "":
                messagebox.showwarning('+ Abrir Chamado: Erro', 'Todos os campos com ( * ) devem ser preenchidos.', parent=root2)
            elif nivel_acesso == 1 and entry_solicitante.get() == "":
                messagebox.showwarning('+ Abrir Chamado: Erro', 'Todos os campos com ( * ) devem ser preenchidos.',
                                       parent=root2)
            else:
                nome_usuario = entryusuario.get().upper()
                solicitante = entry_solicitante.get().upper()
                data_abertura = entrydataabertura.get()
                hora_abertura = entryhoraabertura.get()
                email = entryemail.get()
                ocorr = clique_ocorr.get().upper()
                tipo = cliquetipo.get().upper()
                titulo = entrytitulo.get().upper()
                n_anexo = nome_anexo
                global nome_pasta_anexo
                nome_pasta_anexo = None
                descricao_problema = txtdescr.get("1.0", 'end-1c')
                nome_maquina = entrynomemaquina.get().upper()
                ramal = entryramal.get().upper()
                setor = clique_setor.get()
                status = "Aberto"
                ver = cursor.execute("SELECT * FROM dbo.chamados ORDER BY id_chamado DESC")
                result = ver.fetchone()
                if result == None:
                    max = cursor.execute("SELECT MAX(id_chamado) FROM dbo.chamados")
                    max_resultado = max.fetchone()
                    if nome_anexo != None:
                        nome_pasta_anexo = str(int(max_resultado[0]) + 1)
                        subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
                        subprocess.call(r'net use \\192.168.1.19 /user:impressoras gv2K17ADM', shell=True)

                        fsize = os.stat(caminho_anexo)
                        if fsize.st_size > 5242880:
                            messagebox.showwarning('Erro:', 'Anexo superior a 5MB.', parent=root2)
                        else:
                            pasta = r'\\192.168.1.19/helpdesk/anexos/' + nome_pasta_anexo
                            try:
                                if not os.path.exists(pasta):
                                    os.makedirs(pasta)
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao criar o diretório de destino.', parent=root2)
                                return False
                            try:
                                shutil.copy(caminho_anexo, pasta)
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao enviar o anexo para o diretório.',
                                                       parent=root2)
                                return False
                    try:
                        cursor.execute(
                            "INSERT INTO dbo.chamados (nome_usuario, solicitante, data_abertura, hora_abertura, email, ocorrencia, tipo, titulo, descricao_problema, nome_maquina, ramal, setor, status, nome_anexo, nome_pasta_anexo) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (nome_usuario, solicitante, data_abertura, hora_abertura, email, ocorr, tipo, titulo, descricao_problema, nome_maquina, ramal, setor, status, n_anexo, nome_pasta_anexo))
                        cursor.commit()
                    except:
                        messagebox.showwarning('Erro:', 'Erro ao gravar as informações no banco de dados.',parent=root2)
                        return False

                    messagebox.showinfo('+ Abrir Chamado:', 'Chamado aberto com sucesso. Aguarde para ser atendido.',
                                        parent=root2)
                    subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
                    atualizar_lista_principal()
                    root2.destroy()
                else:
                    cursor.execute("declare @newId int select @newId = max(id_chamado) from chamados DBCC CheckIdent('chamados', RESEED, @newId)")
                    max = cursor.execute("SELECT MAX(id_chamado) FROM dbo.chamados")
                    max_resultado = max.fetchone()
                    if nome_anexo != None:
                        nome_pasta_anexo = str(int(max_resultado[0]) + 1)
                        subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
                        subprocess.call(r'net use \\192.168.1.19 /user:impressoras gv2K17ADM', shell=True)

                        fsize = os.stat(caminho_anexo)
                        if fsize.st_size > 5242880:
                            messagebox.showwarning('Erro:', 'Anexo superior a 5MB.', parent=root2)
                        else:
                            pasta = r'\\192.168.1.19/helpdesk/anexos/' + nome_pasta_anexo
                            try:
                                if not os.path.exists(pasta):
                                    os.makedirs(pasta)
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao criar o diretório de destino.', parent=root2)
                                return False
                            try:
                                shutil.copy(caminho_anexo, pasta)
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao enviar o anexo para o diretório.',
                                                       parent=root2)
                                return False
                    try:
                        cursor.execute(
                            "INSERT INTO dbo.chamados (nome_usuario, solicitante, data_abertura, hora_abertura, email, ocorrencia, tipo, titulo, descricao_problema, nome_maquina, ramal, setor, status, nome_anexo, nome_pasta_anexo) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (nome_usuario, solicitante, data_abertura, hora_abertura, email, ocorr, tipo, titulo, descricao_problema, nome_maquina, ramal, setor, status, n_anexo, nome_pasta_anexo))
                        cursor.commit()
                    except:
                        messagebox.showwarning('Erro:', 'Erro ao gravar as informações no banco de dados.',parent=root2)
                        return False

                    messagebox.showinfo('+ Abrir Chamado:', 'Chamado aberto com sucesso. Aguarde para ser atendido.',
                                        parent=root2)
                    subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
                    atualizar_lista_principal()
                    root2.destroy()

        def solicitante():
            root3 = Toplevel()
            root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
            root3.unbind_class("Button", "<Key-space>")
            root3.focus_force()
            root3.grab_set()
            def busca():
                busca = ent_busca_solic.get().capitalize()
                if busca == "":
                    messagebox.showwarning('Atenção:', 'Campo de solicitante vazio. ', parent=root3)
                else:
                    matching = [s for s in lista if busca in s]
                    tree_principal.delete(*tree_principal.get_children())
                    cont = 0
                    for row in matching:
                        # tree_principal.insert('', 'end', text=" ",values=(row))
                        if cont % 2 == 0:
                            tree_principal.insert('', 'end', values=(row,), tags=('par',))
                        else:
                            tree_principal.insert('', 'end', values=(row,), tags=('impar',))
                        cont += 1
            def busca_bind(event):
                busca = ent_busca_solic.get().capitalize()
                if busca == "":
                    messagebox.showwarning('Atenção:', 'Campo de solicitante vazio. ', parent=root3)
                else:
                    matching = [s for s in lista if busca in s]
                    tree_principal.delete(*tree_principal.get_children())
                    cont = 0
                    for row in matching:
                        # tree_principal.insert('', 'end', text=" ",values=(row))
                        if cont % 2 == 0:
                            tree_principal.insert('', 'end', values=(row,), tags=('par',))
                        else:
                            tree_principal.insert('', 'end', values=(row,), tags=('impar',))
                        cont += 1
            def adicionar_solicitante():
                chamado_select = tree_principal.focus()
                if chamado_select == "":
                    messagebox.showwarning('Atenção:', 'Selecione um nome na lista. ', parent=root3)
                else:
                    nome_solicitante = tree_principal.item(chamado_select, "values")[0]
                    conn.search('DC=gvdobrasil,DC=local', "(&(objectClass=person)(displayName=" + nome_solicitante + "))", SUBTREE, attributes=['sAMAccountName', 'displayName', 'mail'])
                    for i in conn.entries:
                        result = '{0} {1} {2}'.format(i.sAMAccountName.values, i.displayName.values, i.mail.values)
                        conta_nome = format(i.sAMAccountName.values[0])
                        email = format(i.mail.values[0])
                        entry_solicitante.config(state='normal')
                        entry_solicitante.delete(0, END)
                        entry_solicitante.insert(0, conta_nome)
                        entry_solicitante.config(state='disabled')
                        entryemail.config(state='normal')
                        entryemail.delete(0, END)
                        entryemail.insert(0, email)
                        entryemail.config(state='disabled')
                        root3.destroy()
            def duplo_adicionar_solicitante(event):
                adicionar_solicitante()
            def adicionar_solicitante_manual():
                manual = ent_manual.get()
                email_manual = ent_email.get()
                if manual == "" or email_manual == "":
                    messagebox.showwarning('Atenção:', 'Todos os campos devem estar preenchidos.', parent=root3)
                else:
                    entry_solicitante.config(state='normal')
                    entry_solicitante.delete(0, END)
                    entry_solicitante.insert(0, manual)
                    entry_solicitante.config(state='disabled')
                    entryemail.config(state='normal')
                    entryemail.delete(0, END)
                    entryemail.insert(0, email_manual)
                    entryemail.config(state='disabled')
                    root3.destroy()
            def adicionar_solicitante_manual_bind(event):
                manual = ent_manual.get()
                email_manual = ent_email.get()
                if manual == "" or email_manual == "":
                    messagebox.showwarning('Atenção:', 'Todos os campos devem estar preenchidos.', parent=root3)
                else:
                    entry_solicitante.config(state='normal')
                    entry_solicitante.delete(0, END)
                    entry_solicitante.insert(0, manual)
                    entry_solicitante.config(state='disabled')
                    entryemail.config(state='normal')
                    entryemail.delete(0, END)
                    entryemail.insert(0, email_manual)
                    entryemail.config(state='disabled')
                    root3.destroy()

            frame0 = Frame(root3, bg='#ffffff')
            frame0.grid(row=0, column=0, stick='nsew')
            root3.grid_rowconfigure(0, weight=1)
            root3.grid_columnconfigure(0, weight=1)
            frame1 = Frame(frame0, bg="#1d366c")
            frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
            frame2 = Frame(frame0, bg='#ffffff')
            frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=10)
            frame3 = Frame(frame0, bg='#ffffff', padx=10)
            frame3.pack(side=TOP, fill=X, expand=False, anchor='center')
            frame4 = Frame(frame0, bg='#ffffff')
            frame4.pack(side=TOP, fill=X, expand=True, anchor='n')
            frame5 = Frame(frame0, bg='#1d366c', pady=10)
            frame5.pack(side=TOP, fill=X, expand=True, anchor='n')
            frame6 = Frame(frame0, bg='#ffffff')
            frame6.pack(side=TOP, fill=X, expand=True, anchor='n')
            frame7 = Frame(frame0, bg='#1d366c')
            frame7.pack(side=TOP, fill=X, expand=True, anchor='n')


            image_trocasenha2 = Image.open('imagens\\buscaad2.png')
            resize_trocasenha2 = image_trocasenha2.resize((35, 35))
            nova_image_trocasenha2 = ImageTk.PhotoImage(resize_trocasenha2)

            lbllogin = Label(frame1, image=nova_image_trocasenha2, text=" Solicitante", compound="left",
                             bg='#1d366c',
                             fg='#FFFFFF', font=fonte_titulos)
            lbllogin.photo = nova_image_trocasenha2
            lbllogin.grid(row=0, column=1)
            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(2, weight=1)

            ent_busca_solic = Entry(frame2, width=30, font=fonte_padrao, justify='center')
            ent_busca_solic.grid(row=0, column=1, padx=10)
            ent_busca_solic.focus_force()
            ent_busca_solic.bind('<Return>', busca_bind)

            def muda_busca(e):
                image_busca = Image.open('imagens\\lupa_solic_over.png')
                resize_busca = image_busca.resize((25, 25))
                nova_image_busca = ImageTk.PhotoImage(resize_busca)
                btn_busca.photo = nova_image_busca
                btn_busca.config(image=nova_image_busca, fg='#7c7c7c')
            def volta_busca(e):
                image_busca = Image.open('imagens\\lupa_solic.png')
                resize_busca = image_busca.resize((25, 25))
                nova_image_busca = ImageTk.PhotoImage(resize_busca)
                btn_busca.photo = nova_image_busca
                btn_busca.config(image=nova_image_busca, fg='#ffffff')

            image_busca = Image.open('imagens\\lupa_solic.png')
            resize_busca = image_busca.resize((25, 25))
            nova_image_busca = ImageTk.PhotoImage(resize_busca)
            btn_busca = Button(frame2, image=nova_image_busca, bg="#ffffff", fg='#FFFFFF', command=busca,
                               borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btn_busca.photo = nova_image_busca
            btn_busca.grid(row=0, column=2)
            btn_busca.bind("<Enter>", muda_busca)
            btn_busca.bind("<Leave>", volta_busca)
            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(3, weight=1)

            style = ttk.Style()
            # style.theme_use('default')
            style.configure('Treeview',
                            background='#ffffff',
                            rowheight=24,
                            fieldbackground='#ffffff',
                            font=fonte_padrao)
            style.configure("Treeview.Heading",
                            foreground='#000000',
                            background="#ffffff",
                            font=fonte_padrao)
            style.map('Treeview', background=[('selected', '#1d366c')])

            tree_principal = ttk.Treeview(frame3, selectmode='browse')
            vsb = ttk.Scrollbar(frame3, orient="vertical", command=tree_principal.yview)
            vsb.pack(side=RIGHT, fill='y')
            tree_principal.configure(yscrollcommand=vsb.set)
            #vsbx = ttk.Scrollbar(frame3, orient="horizontal", command=tree_principal.xview)
            #vsbx.pack(side=BOTTOM, fill='x')
            #tree_principal.configure(xscrollcommand=vsbx.set)
            tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
            tree_principal["columns"] = ("1")
            tree_principal['show'] = 'headings'
            tree_principal.column("1", width=300, anchor='c')
            tree_principal.heading("1", text="Nome do Solicitante")

            tree_principal.tag_configure('par', background='#e9e9e9')
            tree_principal.tag_configure('impar', background='#ffffff')
            tree_principal.bind("<Double-1>", duplo_adicionar_solicitante)
            frame3.grid_columnconfigure(0, weight=1)
            frame3.grid_columnconfigure(3, weight=1)

            bt2 = Button(frame4, text='Adicionar Solicitante', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                         activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE, command=adicionar_solicitante,
                         font=fonte_padrao, cursor="hand2")
            bt2.grid(row=0, column=2, pady=5, padx=5)


            frame4.grid_columnconfigure(0, weight=1)
            frame4.grid_columnconfigure(3, weight=1)

            conn = Connection("192.168.1.20", "gvdobrasil\\impressoras", "gv2K17ADM", auto_bind=True)
            conn.search('DC=gvdobrasil,DC=local',
                        "(&(objectClass=person)(objectClass=user)(sAMAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))",
                        SUBTREE, attributes=['sAMAccountName', 'displayName'])
            lista = []
            for i in conn.entries:
                #result = '{0} {1}'.format(i.sAMAccountName.values, i.displayName.values)
                conta_nome = format(i.sAMAccountName.values[0])
                display_nome = format(i.displayName.values)
                limpa = display_nome.translate({ord(i): None for i in "['],"})
                lista.append((limpa))

            Label(frame6, text="Solicitante", font=fonte_padrao).grid(row=0, column=1, pady=(10,0))
            ent_manual = Entry(frame6, width=20, font=fonte_padrao, justify='center')
            ent_manual.bind('<Return>', adicionar_solicitante_manual_bind)
            ent_manual.grid(row=1, column=1, padx=6)

            Label(frame6, text="E-mail", font=fonte_padrao).grid(row=0, column=2, pady=(10,0))
            ent_email = Entry(frame6, width=32, font=fonte_padrao, justify='center')
            ent_email.bind('<Return>', adicionar_solicitante_manual_bind)
            ent_email.grid(row=1, column=2, padx=6)

            bt2 = Button(frame6, text='Adicionar Manualmente', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                         activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE, command=adicionar_solicitante_manual,
                         font=fonte_padrao, cursor="hand2")
            bt2.grid(row=2, column=1, pady=10, columnspan=2)

            frame6.grid_columnconfigure(0, weight=1)
            frame6.grid_columnconfigure(3, weight=1)

            Label(frame7, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1, sticky="ew")
            frame7.grid_columnconfigure(0, weight=1)
            frame7.grid_columnconfigure(2, weight=1)
            '''root3.update()
            largura = frame0.winfo_width()
            altura = frame0.winfo_height()
            print(largura, altura)'''
            window_width = 496
            window_height = 518
            screen_width = root2.winfo_screenwidth()
            screen_height = root2.winfo_screenheight()
            x_cordinate = int((screen_width / 2) - (window_width / 2))
            y_cordinate = int((screen_height / 2) - (window_height / 2))
            root3.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
            root3.resizable(0, 0)
            root3.iconbitmap('imagens\\ico.ico')
            root3.title(titulo_todos)

        def cancelar():
            root2.destroy()

        frame0 = Frame(root2, bg='#ffffff')
        frame0.grid(row=0, column=0, stick='nsew')
        root2.grid_rowconfigure(0, weight=1)
        root2.grid_columnconfigure(0, weight=1)
        frame1 = Frame(frame0, bg="#1d366c")
        frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
        frame2 = Frame(frame0, bg='#ffffff')
        frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
        frame3 = Frame(frame0, bg='#ffffff')
        frame3.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
        frame4 = Frame(frame0, bg='#ffffff')
        frame4.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
        frame5 = Frame(frame0, bg='#ffffff')
        frame5.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
        frame6 = Frame(frame0, bg='#1d366c') #linha
        frame6.pack(side=TOP, fill=X, expand=False, anchor='center')
        frame7 = Frame(frame0, bg='#ffffff')
        frame7.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
        frame8 = Frame(frame0, bg='#ffffff')
        frame8.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
        frame9 = Frame(frame0, bg='#ffffff')
        frame9.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
        frame10 = Frame(frame0, bg='#1d366c')
        frame10.pack(side=TOP, fill=X, expand=False, anchor='center')

        Label(frame1, image=nova_image_chamado, text=" + Abrir Chamado", compound="left", bg='#1d366c', fg='#FFFFFF',
              font=fonte_titulos).grid(row=0, column=1)
        frame1.grid_columnconfigure(0, weight=1)
        frame1.grid_columnconfigure(2, weight=1)

        Label(frame2, text="Usuário Logado:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w")
        entryusuario = Entry(frame2, font=fonte_padrao, justify='center')
        entryusuario.grid(row=1, column=1, sticky="w")
        entryusuario.insert(0, usuariologado)
        entryusuario.config(state='disabled')



        def muda_solicitante(e):
            image_solicitante = Image.open('imagens\\buscaad_over.png')
            resize_solicitante = image_solicitante.resize((20, 20))
            nova_image_solicitante = ImageTk.PhotoImage(resize_solicitante)
            btnsolicitante.photo = nova_image_solicitante
            btnsolicitante.config(image=nova_image_solicitante, fg='#7c7c7c')
        def volta_solicitante(e):
            image_solicitante = Image.open('imagens\\buscaad.png')
            resize_solicitante = image_solicitante.resize((20, 20))
            nova_image_solicitante = ImageTk.PhotoImage(resize_solicitante)
            btnsolicitante.photo = nova_image_solicitante
            btnsolicitante.config(image=nova_image_solicitante, fg='#8B0000')

        image_solicitante = Image.open('imagens\\buscaad.png')
        resize_solicitante = image_solicitante.resize((20, 20))
        nova_image_solicitante = ImageTk.PhotoImage(resize_solicitante)
        btnsolicitante = Button(frame2, image=nova_image_solicitante, text=" Solicitante.", compound="left",
                          font=fonte_padrao, bg='#ffffff', fg='#8B0000', command=solicitante,
                          borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
        btnsolicitante.photo = nova_image_solicitante
        btnsolicitante.grid(row=0, column=2, sticky="w", padx=12)
        btnsolicitante.bind("<Enter>", muda_solicitante)
        btnsolicitante.bind("<Leave>", volta_solicitante)
        entry_solicitante = Entry(frame2, font=fonte_padrao, justify='center')
        entry_solicitante.grid(row=1, column=2, sticky="w", padx=12)
        if nivel_acesso == 0:
            entry_solicitante.insert(0, usuariologado)
            entry_solicitante.config(state='disabled')
            btnsolicitante.config(state='disabled')
            try:
                conn = Connection("192.168.1.20", "gvdobrasil\\impressoras", "gv2K17ADM", auto_bind=True)
                conn.search('DC=gvdobrasil,DC=local', "(&(objectClass=person)(sAMAccountName=" + usuariologado + "))",
                        SUBTREE, attributes=['sAMAccountName', 'displayName', 'mail'])
                for i in conn.entries:
                    result = '{0} {1} {2}'.format(i.sAMAccountName.values, i.displayName.values, i.mail.values)
                    conta_nome = format(i.sAMAccountName.values[0])
                    email = format(i.mail.values[0])
            except:
                messagebox.showerror('Erro:', 'Erro ao adicionar o e-mail.', parent=root2)

        else:
            entry_solicitante.config(state='disabled')

        Label(frame2, text="E-mail:", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0, column=3, sticky="w")
        entryemail = Entry(frame2, font=fonte_padrao, justify='center', width=47)
        entryemail.grid(row=1, column=3, sticky="w")
        if nivel_acesso == 0:
            entryemail.insert(0, email)
        entryemail.config(state='disabled')

        frame2.grid_columnconfigure(0, weight=1)
        frame2.grid_columnconfigure(5, weight=1)

        Label(frame3, text="Ocorrência: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0, column=1, sticky="w")
        clique_ocorr = StringVar() 
        drop_ocorr = OptionMenu(frame3, clique_ocorr, "Acessos", "Câmeras", "E-mail", "Hardware", "Impressora", "Internet", "Protheus", "Rádio", "Rede", "Softwares", "Telefonia", "VPN", "Windows",
                              command=dropselecao_ocorr)
        drop_ocorr.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                        highlightthickness=0, relief=RIDGE, width=21, cursor="hand2")
        drop_ocorr.grid(row=1, column=1, sticky="w")


        Label(frame3, text="Tipo: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0, column=2, sticky="w", padx=10)
        cliquetipo = StringVar()
        droptipo = OptionMenu(frame3, cliquetipo, "Solicitação", "Problemas", "Dúvidas", "Melhorias", "Projetos", command=dropselecaotipo)
        droptipo.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                        highlightthickness=0, relief=RIDGE, width=17, cursor="hand2")
        droptipo.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=17)
        droptipo.grid(row=1, column=2, sticky="w", padx=10)
        Label(frame3, text="Título Predefinido: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0, column=3, sticky="w")
        cliquesub = StringVar()
        global dropsub
        dropsub = OptionMenu(frame3, cliquesub, "")
        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=46, cursor="hand2")
        dropsub.grid(row=1, column=3, sticky="w")
        frame3.grid_columnconfigure(0, weight=1)
        frame3.grid_columnconfigure(4, weight=1)

        lbltitulo = Label(frame4, text="Título: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000')
        lbltitulo.grid(row=0, column=1, sticky="w")
        entrytitulo = Entry(frame4, font=fonte_padrao, justify='center', width=64)
        entrytitulo.grid(row=1, column=1, sticky="ew", padx=(0,9))
        entrytitulo.config(state=DISABLED)

        image_anexo = Image.open('imagens\\anexo.png')
        resize_anexo = image_anexo.resize((15, 20))
        nova_image_anexo = ImageTk.PhotoImage(resize_anexo)
        btnanexo = Button(frame4, image=nova_image_anexo, text=" Anexar arquivo.", compound="left",
                          font=fonte_padrao, bg='#ffffff', fg='#1d366c', command=anexo,
                          borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
        btnanexo.photo = nova_image_anexo
        btnanexo.grid(row=0, column=2, sticky="w", padx=(9,0))

        entryanexo = Entry(frame4, font=fonte_padrao, justify='center', width=25)
        entryanexo.grid(row=1, column=2, sticky="ew", padx=(9,0))
        entryanexo.config(state='disabled')
        frame4.grid_columnconfigure(0, weight=1)
        frame4.grid_columnconfigure(3, weight=1)

        Label(frame5, text="Descrição do Problema: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0, column=1, sticky="w")
        txtdescr = scrolledtext.ScrolledText(frame5, width=90, height=10, font=fonte_padrao, wrap=WORD)
        txtdescr.grid(row=1, column=1)
        frame5.grid_columnconfigure(0, weight=1)
        frame5.grid_columnconfigure(2, weight=1)

        #FRAME6 LINHA

        Label(frame7, text="Data de Abertura:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w", padx=12)
        entrydataabertura = Entry(frame7, font=fonte_padrao, justify='center')
        entrydataabertura.grid(row=1, column=1, sticky="w", padx=12)
        entrydataabertura.insert(0, data)
        entrydataabertura.config(state='disabled')

        Label(frame7, text="Hora de Abertura:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w", padx=12)
        entryhoraabertura = Entry(frame7, font=fonte_padrao, justify='center')
        entryhoraabertura.grid(row=1, column=2, sticky="w", padx=12)
        entryhoraabertura.insert(0, hora)
        entryhoraabertura.config(state='disabled')

        frame7.grid_columnconfigure(0, weight=1)
        frame7.grid_columnconfigure(3, weight=1)

        Label(frame8, text="Nome da Máquina:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=1, column=3, sticky="w")
        entrynomemaquina = Entry(frame8, font=fonte_padrao, justify='center')
        entrynomemaquina.grid(row=2, column=3, sticky="w")

        Label(frame8, text="Ramal:", font=fonte_padrao, bg='#ffffff').grid(row=1, column=4, sticky="w", padx=14)
        entryramal = Entry(frame8, font=fonte_padrao, justify='center')
        entryramal.grid(row=2, column=4, sticky="w", padx=14)

        Label(frame8, text="Setor: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=1, column=5, sticky="w")

        OptionList = [
            "Aciaria",
            "Almoxarifado",
            "Ambulatório",
            "Auditoria",
            "Balança",
            "Comercial",
            "Compras",
            "Contabilidade",
            "Custos",
            "Diretoria",
            "EHS",
            "Elétrica",
            "Engenharia",
            "Faturamento",
            "Financeiro",
            "Fiscal",
            "Lab Inspeção",
            "Lab Mecânico",
            "Lab Químico",
            "Laminação",
            "Logística",
            "Oficina de Cilindros",
            "Oficina Mecânica",
            "Pátio de Sucata",
            "PCP",
            "Planta D´agua",
            "Planta de Escória",
            "Portaria",
            "Qualidade",
            "Refratários",
            "Refrigeração",
            "RH",
            "Subestação",
            "TI",
            "Utilidades"
        ]
        clique_setor = StringVar()
        drop_setor = OptionMenu(frame8, clique_setor, *OptionList)
        drop_setor.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                        highlightthickness=0, relief=RIDGE, width=26, cursor="hand2")
        drop_setor.grid(row=2, column=5, sticky="w")
        #entrysetor = Entry(frame6, font=fonte_padrao, justify='center')
        #entrysetor.grid(row=1, column=3, sticky="w")

        frame8.grid_columnconfigure(0, weight=1)
        frame8.grid_columnconfigure(6, weight=1)

        bt1 = Button(frame9, text='Salvar', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                     activeforeground="#FFFFFF", highlightthickness=0, width=10, relief=RIDGE, command=salvar,
                     font=fonte_padrao, cursor="hand2")
        bt1.grid(row=0, column=1, padx=5)
        bt2 = Button(frame9, text='Cancelar', width=10, relief=RIDGE, command=cancelar, font=fonte_padrao, cursor="hand2")
        bt2.grid(row=0, column=2, padx=5)
        frame9.grid_columnconfigure(0, weight=1)
        frame9.grid_columnconfigure(3, weight=1)

        Label(frame10, text=" ", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos).grid(row=0, column=1)
        frame10.grid_columnconfigure(0, weight=1)
        frame10.grid_columnconfigure(2, weight=1)

        '''root2.update()
        largura = frame0.winfo_width()
        altura = frame0.winfo_height()
        print(largura, altura)'''
        window_width = 670
        window_height = 639
        screen_width = root2.winfo_screenwidth()
        screen_height = root2.winfo_screenheight() - 70
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
        #root2.resizable(0, 0)
        root2.configure(bg='#000000')
        root2.iconbitmap('imagens\\ico.ico')

    # /////////////////////////////FIM ABRIR CHAMADO/////////////////////////////

    # /////////////////////////////ATENDIMENTO/////////////////////////////
    def atendimento_bind(event):
        if nivel_acesso == 0:
            messagebox.showwarning('Atenção:', 'Módulo (Atendimento) bloqueado.', parent=root)
        else:
            atendimento()
    def atendimento():

        def layout():
            root2 = Toplevel(root)
            root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
            root2.unbind_class("Button", "<Key-space>")
            root2.focus_force()
            root2.grab_set()

            def cancelar():
                root2.destroy()
            def salvar():
                nome_analista = entryanalista2.get()
                prioridade = clique_prioridade.get()
                status = clique_status.get()
                data_atentimento = entrydataatendimento2.get()
                data_encerramento = entrydataencerramento.get()
                solucao = txtsolucao.get("1.0", 'end-1c')
                # txtinteracao.config(state='normal')
                interacao = txtinteracao.get("1.0", 'end-1c')
                chamado_enviado = entry_enviar_chamado.get()
                tupla = (nome_analista, status, data_atentimento, data_encerramento, solucao, n_chamado)
                print(chamado_enviado)

                if status == "Aberto":
                    messagebox.showwarning('Atenção !',
                                           'Para concluir a alteração, o campo "STATUS" deve ser alterado.',
                                           parent=root2)
                elif status == "Em andamento" and data_atentimento == '':
                    messagebox.showwarning('Atenção !',
                                           'Para concluir a alteração, o campo "DATA DO ATENDIMENTO" deve estar preenchido.',
                                           parent=root2)
                elif status == "Em andamento" and data_encerramento != '':
                    messagebox.showwarning('Atenção !',
                                           'O campo "DATA DE ENCERRAMENTO" NÃO deve estar preenchido caso o status seja "EM ANDAMENTO".',
                                           parent=root2)
                    entrydataencerramento.config(state='normal')
                    entrydataencerramento.delete(0, END)
                    entrydataencerramento.config(state='disabled')
                elif status == "Encerrado" and data_encerramento == '':
                    messagebox.showwarning('Atenção !',
                                           'Para finalizar o chamado, o campo "DATA DE ENCERRAMENTO" deve estar preenchido.',
                                           parent=root2)
                elif status == "Encerrado" and solucao == '':
                    messagebox.showwarning('Atenção !',
                                           'Para finalizar o chamado, o campo "SOLUÇÃO" deve estar preenchido.',
                                           parent=root2)
                elif prioridade == "":
                    messagebox.showwarning('Atenção !', 'Defina uma prioridade para o chamado.',
                                           parent=root2)
                elif solucao != "" and data_encerramento == "":
                    messagebox.showwarning('Atenção !',
                                           'O campo "Solução" só deve ser utilizado para finalizar o chamado. Utilize o campo de "Interação".',
                                           parent=root2)
                elif status == "Encerrado" and solucao != '' and data_encerramento != '':
                    data_comparacao_atendimento = time.strptime(data_atentimento, "%d/%m/%Y")
                    data_comparacao_encerramento = time.strptime(data_encerramento, "%d/%m/%Y")
                    data_atual = time.strptime(data, "%d/%m/%Y")
                    if data_comparacao_atendimento > data_comparacao_encerramento:
                        messagebox.showwarning('Atenção:',
                                               'A "Data de Encerramento" não deve ser menor que a "Data de Atendimento"',
                                               parent=root2)
                    elif data_comparacao_encerramento > data_atual:
                        messagebox.showwarning('Atenção:',
                                               'A "Data de Encerramento" não deve ser superior a data de hoje.',
                                               parent=root2)
                    else:
                        cursor.execute(
                            "UPDATE helpdesk.dbo.chamados SET id_analista = ?, prioridade = ?, status = ?, data_atendimento = ?, data_encerramento = ?, interacao = ?, resolucao = ?, enviar_chamado = ? WHERE id_chamado = ?",
                            (nome_analista, prioridade, status, data_atentimento, data_encerramento, interacao, solucao,
                             chamado_enviado,
                             n_chamado))
                        cursor.commit()
                        messagebox.showinfo('Atendimento:', 'Alteração realizada com sucesso!', parent=root2)
                        atualizar_lista_principal_encerrado()
                        #////////////ENVIA E-MAIL FINALIZANDO CHAMADO
                        sender_email = "naoresponder@gruposimec.com.br"
                        receiver_email = result[21]
                        password = "Qu@@258147"
                        solucao_html = solucao.replace("\n", "<br>")
                        

                        message = MIMEMultipart("alternative")
                        message["Subject"] = "HelpDesk - Chamado nº" + str(result[0]) + " - (Não responder)"
                        message["From"] = sender_email
                        message["To"] = receiver_email

                        # Create the plain-text and HTML version of your message
                        text = """\
                        Seu chamado nº """ + str(result[0]) + """ foi encerrado por """ + str(nome_analista) + """.

                        Este e-mail não precisa ser respondido."""

                        html = """\
                        <html>
                          <body>
                         <center>	
                        <font size="2" face="Arial" >
                        <table width=100% border=0>
                          <tr style="background-color:#01336e">
                            <td align=center><p style= "font-family:Arial; font-size:50px; color:white"><b>HelpDesk</b></p></td>
                          </tr>
                          <tr>
                            <td align=center>Seu chamado <b>nº """ + str(
                            result[0]) + """</b> foi finalizado por <b>""" + str(nome_analista) + """</b>. </td>
                          </tr>
                          <tr>
                            <td align=center style="color:#01336e"><b>Solução:<br>"""+str(solucao_html)+ """</b></td>
                          </tr>
                          <tr>
                            <td align=center><br><br><b>Obs: Este e-mail não precisa ser respondido.<b></td>
                          </tr>
                        </table>
                        </font>
                        </center>	

                          </body>
                        </html>

                        """

                        # Turn these into plain/html MIMEText objects
                        part1 = MIMEText(text, "plain")
                        part2 = MIMEText(html, "html")

                        # Add HTML/plain-text parts to MIMEMultipart message
                        # The email client will try to render the last part first
                        message.attach(part1)
                        message.attach(part2)

                        # Create secure connection with server and send email
                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL("smtps.uhserver.com", 465, context=context) as server:
                            server.login(sender_email, password)
                            server.sendmail(
                                sender_email, receiver_email, message.as_string()
                            )
                        root2.destroy()
                elif status == "Cancelado" and solucao != '' and data_encerramento != '':
                    data_comparacao_atendimento = time.strptime(data_atentimento, "%d/%m/%Y")
                    data_comparacao_encerramento = time.strptime(data_encerramento, "%d/%m/%Y")
                    data_atual = time.strptime(data, "%d/%m/%Y")
                    if data_comparacao_atendimento > data_comparacao_encerramento:
                        messagebox.showwarning('Atenção:',
                                               'A "Data de Encerramento" não deve ser menor que a "Data de Atendimento"',
                                               parent=root2)
                    elif data_comparacao_encerramento > data_atual:
                        messagebox.showwarning('Atenção:',
                                               'A "Data de Encerramento" não deve ser superior a data de hoje.',
                                               parent=root2)
                    else:
                        cursor.execute(
                            "UPDATE helpdesk.dbo.chamados SET id_analista = ?, prioridade = ?, status = ?, data_atendimento = ?, data_encerramento = ?, interacao = ?, resolucao = ?, enviar_chamado = ? WHERE id_chamado = ?",
                            (nome_analista, prioridade, status, data_atentimento, data_encerramento, interacao, solucao,
                             chamado_enviado,
                             n_chamado))
                        cursor.commit()
                        messagebox.showinfo('Atendimento:', 'Alteração realizada com sucesso!', parent=root2)
                        atualizar_lista_principal_encerrado()

                        sender_email = "naoresponder@gruposimec.com.br"
                        receiver_email = result[21]
                        password = "Qu@@258147"
                        solucao_html = solucao.replace("\n", "<br>")
                        

                        message = MIMEMultipart("alternative")
                        message["Subject"] = "HelpDesk - Chamado nº" + str(result[0]) + " - (Não responder)"
                        message["From"] = sender_email
                        message["To"] = receiver_email

                        # Create the plain-text and HTML version of your message
                        text = """\
                        Seu chamado nº """ + str(result[0]) + """ foi encerrado por """ + str(nome_analista) + """.

                        Este e-mail não precisa ser respondido."""

                        html = """\
                        <html>
                          <body>
                         <center>	
                        <font size="2" face="Arial" >
                        <table width=100% border=0>
                          <tr style="background-color:#01336e">
                            <td align=center><p style= "font-family:Arial; font-size:50px; color:white"><b>HelpDesk</b></p></td>
                          </tr>
                          <tr>
                            <td align=center>Seu chamado <b>nº """ + str(
                            result[0]) + """</b> foi finalizado por <b>""" + str(nome_analista) + """</b>. </td>
                          </tr>
                          <tr>
                            <td align=center style="color:#01336e"><b>Solução:<br>"""+str(solucao_html)+ """</b></td>
                          </tr>
                          <tr>
                            <td  align=center><br><br><b>Obs: Este e-mail não precisa ser respondido.<b></td>
                          </tr>
                        </table>
                        </font>
                        </center>	

                          </body>
                        </html>

                        """

                        # Turn these into plain/html MIMEText objects
                        part1 = MIMEText(text, "plain")
                        part2 = MIMEText(html, "html")

                        # Add HTML/plain-text parts to MIMEMultipart message
                        # The email client will try to render the last part first
                        message.attach(part1)
                        message.attach(part2)

                        # Create secure connection with server and send email
                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL("smtps.uhserver.com", 465, context=context) as server:
                            server.login(sender_email, password)
                            server.sendmail(
                                sender_email, receiver_email, message.as_string()
                            )

                        root2.destroy()
                else:
                    cursor.execute(
                        "UPDATE helpdesk.dbo.chamados SET id_analista = ?, prioridade = ?, status = ?, data_atendimento = ?, data_encerramento = ?, interacao = ?, resolucao = ?, enviar_chamado = ? WHERE id_chamado = ?",
                        (nome_analista, prioridade, status, data_atentimento, data_encerramento, interacao, solucao,
                         chamado_enviado, n_chamado))
                    cursor.commit()
                    messagebox.showinfo('Atendimento:', 'Alteração realizada com sucesso!', parent=root2)
                    atualizar_lista_principal()
                    if result[13] == None and result[21] != None:
                        sender_email = "naoresponder@gruposimec.com.br"
                        receiver_email = result[21]
                        password = "Qu@@258147"
                        

                        message = MIMEMultipart("alternative")
                        message["Subject"] = "HelpDesk - Chamado nº" + str(result[0]) + " - (Não responder)"
                        message["From"] = sender_email
                        message["To"] = receiver_email

                        # Create the plain-text and HTML version of your message
                        text = """\
                        Seu chamado nº """ + str(result[0]) + """ será atendido por """ + str(nome_analista) + """.

                        Qualquer dúvida, usar o campo de Interação do nosso sistema.

                        Este e-mail não precisa ser respondido."""

                        html = """\
                        <html>
                          <body>
                         <center>	
                        <font size="2" face="Arial" >
                        <table width=100% border=0>
                          <tr style="background-color:#01336e">
                            <td align=center><p style= "font-family:Arial; font-size:50px; color:white"><b>HelpDesk</b></p></td>
                          </tr>
                          <tr>
                            <td align=center>Seu chamado <b>nº """ + str(
                            result[0]) + """</b> será atendido por  <b>""" + str(nome_analista) + """</b>. </td>
                          </tr>
                          <tr>
                            <td  align=center><br>Qualquer dúvida , usar o campo de "Interação" do nosso sistema.</td>
                          </tr>
                          <tr>
                            <td  align=center><br><br><b>Obs: Este e-mail não precisa ser respondido.<b></td>
                          </tr>
                        </table>
                        </font>
                        </center>	

                          </body>
                        </html>

                        """

                        # Turn these into plain/html MIMEText objects
                        part1 = MIMEText(text, "plain")
                        part2 = MIMEText(html, "html")

                        # Add HTML/plain-text parts to MIMEMultipart message
                        # The email client will try to render the last part first
                        message.attach(part1)
                        message.attach(part2)

                        # Create secure connection with server and send email
                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL("smtps.uhserver.com", 465, context=context) as server:
                            server.login(sender_email, password)
                            server.sendmail(
                                sender_email, receiver_email, message.as_string()
                            )
                    
                    if email_interacao == 1:
                        sender_email = "naoresponder@gruposimec.com.br"
                        receiver_email = result[21]
                        password = "Qu@@258147"
                        interacao_html = interacao.replace("\n", "<br>")

                        message = MIMEMultipart("alternative")
                        message["Subject"] = "HelpDesk - Interação do Chamado nº" + str(
                            result[0]) + " - (Não responder)"
                        message["From"] = sender_email
                        message["To"] = receiver_email

                        # Create the plain-text and HTML version of your message
                        text = """\
                        O analista  """ + str(nome_analista) + """ interagiu ao seu chamado de nº """ + str(result[0]) + """.

                        Acesse o sistema o mais breve possível e responda a interação.

                        Obs: Caso não haja novas interações, encerraremos o chamado."""

                        html = """\
                        <html>
                          <body>
                         <center>	
                        <font size="2" face="Arial" >
                        <table width=100% border=0>
                          <tr style="background-color:#01336e">
                            <td align=center><p style= "font-family:Arial; font-size:50px; color:white"><b>HelpDesk</b></p></td>
                          </tr>
                          <tr>
                            <td align=center>O analista <b>""" + str(
                            nome_analista) + """</b> interagiu ao seu chamado de nº <b>""" + str(result[0]) + """</b> e aguarda sua resposta. </td>
                          </tr>
                          <tr>
                            <td  align=center><br>Acesse o sistema o mais breve possível e responda a interação.</td>
                          </tr>
                          <tr>
                            <td>Histórico da interação<br><br>"""+str(interacao_html)+"""</td>
                          </tr>
                          <tr>
                            <td  align=center><br><br><b>Caso não haja novas interações, encerraremos o chamado.<b></td>
                          </tr>
                        </table>
                        </font>
                        </center>	

                          </body>
                        </html>

                        """

                        # Turn these into plain/html MIMEText objects
                        part1 = MIMEText(text, "plain")
                        part2 = MIMEText(html, "html")

                        # Add HTML/plain-text parts to MIMEMultipart message
                        # The email client will try to render the last part first
                        message.attach(part1)
                        message.attach(part2)

                        # Create secure connection with server and send email
                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL("smtps.uhserver.com", 465, context=context) as server:
                            server.login(sender_email, password)
                            server.sendmail(
                                sender_email, receiver_email, message.as_string()
                            )
                   
                    root2.destroy()
            def setup_entradas():
                def config_entradas():
                    if result[16] == None:
                        entrysolicitante.insert(0, '')
                        entrysolicitante.config(state='disabled')
                    else:
                        entrysolicitante.insert(0, result[16])
                        entrysolicitante.config(state='disabled')
                    entrydataabertura.insert(0, result[2])
                    entrydataabertura.config(state='disabled')
                    if result[3] == None:
                        entryhoraabertura.insert(0, "")
                        entryhoraabertura.config(state='disabled')
                    else:
                        entryhoraabertura.insert(0, result[3])
                        entryhoraabertura.config(state='disabled')
                    if result[21] == None:
                        entryemail.insert(0, '')
                        entryemail.config(state='disabled')
                    else:
                        entryemail.insert(0, result[21])
                        entryemail.config(state='disabled')
                    entrynomemaquina.insert(0, result[8])
                    entrynomemaquina.config(state='disabled')
                    entryramal.insert(0, result[9])
                    entryramal.config(state='disabled')
                    entrysetor.insert(0, result[10])
                    entrysetor.config(state='disabled')
                    if result[17] == None:
                        entryocorrencia.insert(0, '')
                        entryocorrencia.config(state='disabled')
                    else:
                        entryocorrencia.insert(0, result[17])
                        entryocorrencia.config(state='disabled')
                    entrytitulo.insert(0, result[5])
                    entrytitulo.config(state='disabled')
                    entrytipo.insert(0, result[4])
                    entrytipo.config(state='disabled')
                    txtdescr_atendimento.insert(END, result[7])
                    txtdescr_atendimento.config(fg='#696969')
                    txtdescr_atendimento.config(state='disabled')
                    if result[6] == None and result[22] == None:
                        btabrir_anexo.config(state='disabled')
                    entryanalista2.insert(0, usuariologado)
                    entryanalista2.config(state='disabled')
                    if result[18] == None:
                        clique_prioridade.set('')
                    else:
                        clique_prioridade.set(result[18])
                    if result[11] == 'Aberto':
                        clique_status.set('Em andamento')
                        entrydataatendimento2.config(state='normal')
                        entrydataatendimento2.insert(0, data)
                        entrydataatendimento2.config(state='disabled')
                        #drop_redirec.config(state='disabled')

                    else:
                        clique_status.set(result[11])

                    if result[12] != None:
                        entrydataatendimento2.config(state='normal')
                        entrydataatendimento2.insert(0, result[12])
                        entrydataatendimento2.config(state='disabled')
                        # clique_status.set('Em andamento')

                    if result[15] == None:
                        entrydataencerramento.config(state='disabled')
                    elif result[15] != '':
                        entrydataencerramento.config(state='normal')
                        entrydataencerramento.insert(0, result[15])
                        entrydataencerramento.config(state='disabled')
                        clique_status.set('Encerrado')
                        txtsolucao.insert(END, result[14])
                        txtsolucao.config(fg='#696969')
                        txtsolucao.config(state='disabled')
                        btcancelar_atendimento.config(state='disabled')
                        btconfirma_atencimento.config(bg="#C0C0C0", fg="#ffffff", state='disabled')
                        drop_status.config(state='disabled')
                    else:
                        entrydataencerramento.config(state='disabled')
                    if result[19] != None:
                        txtinteracao.config(state='normal')
                        txtinteracao.insert(END, result[19])
                        txtinteracao.config(state='disabled')
                    if result[24] != None:
                        entry_enviar_chamado.config(state='normal')
                        entry_enviar_chamado.insert(END, result[24])
                        entry_enviar_chamado.config(state='disabled')
                if troca_analista == 0:
                    config_entradas()
                else:
                    config_entradas()
                    hora = time.strftime('%H:%M:%S', time.localtime())
                    txtinteracao.config(state='normal')
                    txtinteracao.insert(END,
                                        f'\nTroca de Analista: ({data} - {hora})\n{result[13]} -> {usuariologado}.\n-----------------------------------------------------------------------')
                    txtinteracao.config(state='disabled')


            def abrir_anexo():
                if result[6] != None:
                    def conversorarquivo(data, filename):
                        with open(filename, 'wb') as file:
                            file.write(data)

                    nome_arquivo = result[6]
                    abreviacao_extensao = result[20].upper()
                    nome_extensao = result[20]
                    caminho = filedialog.asksaveasfilename(defaultextension=".*",
                                                           initialfile='Anexo_Chamado (' + str(result[0]) + ')',
                                                           initialdir='os.path.expanduser(default_dir)',
                                                           title='Salvar Anexo..',
                                                           filetypes=[
                                                               (str(abreviacao_extensao), "*" + nome_extensao)], parent=root2)
                    conversorarquivo(nome_arquivo, caminho)

                else:
                    subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
                    subprocess.call(r'net use \\192.168.1.19 /user:impressoras gv2K17ADM', shell=True)

                    def conversorarquivo(origem, destino):
                        try:
                            shutil.copy(origem, destino)
                        except:
                            messagebox.showerror('Erro:',
                                                 'Arquivo de destino não encontrado ou Download cancelado.',
                                                 parent=root2)
                            return False
                        messagebox.showinfo('Download:', 'Download concluído com sucesso', parent=root2)

                    cursor.execute("SELECT * FROM dbo.chamados WHERE nome_anexo=? AND id_chamado=?",
                                   (result[22], result[0],))
                    busca = cursor.fetchone()
                    nome_pasta = result[23]
                    pasta = r'\\192.168.1.19/helpdesk/anexos/' + nome_pasta
                    caminhocompleto = os.path.join(pasta, busca[22])
                    nome = busca[22]
                    destino = filedialog.asksaveasfilename(
                        initialfile=str(nome),
                        initialdir='os.path.expanduser(default_dir)',
                        title='Salvar Anexo..',
                        filetypes=[("Todos os arquivos", "*")], parent=root2)
                    conversorarquivo(caminhocompleto, destino)

                    subprocess.call(r'net use /delete \\192.168.1.19', shell=True)

            def encaminhar():
                root3 = Toplevel()
                root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
                root3.unbind_class("Button", "<Key-space>")
                root3.focus_force()
                root3.grab_set()

                def busca():
                    busca = ent_busca_solic.get().capitalize()
                    if busca == "":
                        messagebox.showwarning('Atenção:', 'Campo de solicitante vazio. ', parent=root3)
                    else:
                        matching = [s for s in lista if busca in s]
                        tree_principal.delete(*tree_principal.get_children())
                        cont = 0
                        for row in matching:
                            # tree_principal.insert('', 'end', text=" ",values=(row))
                            if cont % 2 == 0:
                                tree_principal.insert('', 'end', values=(row,), tags=('par',))
                            else:
                                tree_principal.insert('', 'end', values=(row,), tags=('impar',))
                            cont += 1

                def busca_bind(event):
                    busca = ent_busca_solic.get().capitalize()
                    if busca == "":
                        messagebox.showwarning('Atenção:', 'Campo de solicitante vazio. ', parent=root3)
                    else:
                        matching = [s for s in lista if busca in s]
                        tree_principal.delete(*tree_principal.get_children())
                        cont = 0
                        for row in matching:
                            # tree_principal.insert('', 'end', text=" ",values=(row))
                            if cont % 2 == 0:
                                tree_principal.insert('', 'end', values=(row,), tags=('par',))
                            else:
                                tree_principal.insert('', 'end', values=(row,), tags=('impar',))
                            cont += 1

                def adicionar_solicitante():
                    chamado_select = tree_principal.focus()
                    if chamado_select == "":
                        messagebox.showwarning('Atenção:', 'Selecione um nome na lista. ', parent=root3)
                    else:
                        nome_solicitante = tree_principal.item(chamado_select, "values")[0]
                        conn.search('DC=gvdobrasil,DC=local',
                                    "(&(objectClass=person)(displayName=" + nome_solicitante + "))", SUBTREE,
                                    attributes=['sAMAccountName', 'displayName', 'mail'])
                        for i in conn.entries:
                            result = '{0} {1} {2}'.format(i.sAMAccountName.values, i.displayName.values,
                                                          i.mail.values)
                            conta_nome = format(i.sAMAccountName.values[0])
                            email = format(i.mail.values[0])
                            entry_enviar_chamado.config(state='normal')
                            entry_enviar_chamado.delete(0, END)
                            entry_enviar_chamado.insert(0, conta_nome)
                            entry_enviar_chamado.config(state='disabled')
                            entryemail.config(state='normal')
                            entryemail.delete(0, END)
                            entryemail.insert(0, email)
                            entryemail.config(state='disabled')
                            clique_status.set("Aguardando Usuário")
                            root3.destroy()

                def duplo_adicionar_solicitante(event):
                    adicionar_solicitante()

                def adicionar_solicitante_manual():
                    manual = ent_manual.get()
                    email_manual = ent_email.get()
                    if manual == "" or email_manual == "":
                        messagebox.showwarning('Atenção:', 'Todos os campos devem estar preenchidos.', parent=root3)
                    else:
                        entry_enviar_chamado.config(state='normal')
                        entry_enviar_chamado.delete(0, END)
                        entry_enviar_chamado.insert(0, manual)
                        entry_enviar_chamado.config(state='disabled')
                        entryemail.config(state='normal')
                        entryemail.delete(0, END)
                        entryemail.insert(0, email_manual)
                        entryemail.config(state='disabled')
                        root3.destroy()

                def adicionar_solicitante_manual_bind(event):
                    adicionar_solicitante_manual()

                frame0 = Frame(root3, bg='#ffffff')
                frame0.grid(row=0, column=0, stick='nsew')
                root3.grid_rowconfigure(0, weight=1)
                root3.grid_columnconfigure(0, weight=1)
                frame1 = Frame(frame0, bg="#1d366c")
                frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
                frame2 = Frame(frame0, bg='#ffffff')
                frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=10)
                frame3 = Frame(frame0, bg='#ffffff', padx=10)
                frame3.pack(side=TOP, fill=X, expand=False, anchor='center')
                frame4 = Frame(frame0, bg='#ffffff')
                frame4.pack(side=TOP, fill=X, expand=True, anchor='n')
                frame5 = Frame(frame0, bg='#1d366c', pady=10)
                frame5.pack(side=TOP, fill=X, expand=True, anchor='n')
                frame6 = Frame(frame0, bg='#ffffff')
                frame6.pack(side=TOP, fill=X, expand=True, anchor='n')
                frame7 = Frame(frame0, bg='#1d366c')
                frame7.pack(side=TOP, fill=X, expand=True, anchor='n')

                image_trocasenha2 = Image.open('imagens\\seta2.png')
                resize_trocasenha2 = image_trocasenha2.resize((35, 35))
                nova_image_trocasenha2 = ImageTk.PhotoImage(resize_trocasenha2)

                lbllogin = Label(frame1, image=nova_image_trocasenha2, text=" Enviar Chamado", compound="left",
                                 bg='#1d366c',
                                 fg='#FFFFFF', font=fonte_titulos)
                lbllogin.photo = nova_image_trocasenha2
                lbllogin.grid(row=0, column=1)
                frame1.grid_columnconfigure(0, weight=1)
                frame1.grid_columnconfigure(2, weight=1)

                ent_busca_solic = Entry(frame2, width=30, font=fonte_padrao, justify='center')
                ent_busca_solic.grid(row=0, column=1, padx=10)
                ent_busca_solic.focus_force()
                ent_busca_solic.bind('<Return>', busca_bind)

                def muda_busca(e):
                    image_busca = Image.open('imagens\\lupa_solic_over.png')
                    resize_busca = image_busca.resize((25, 25))
                    nova_image_busca = ImageTk.PhotoImage(resize_busca)
                    btn_busca.photo = nova_image_busca
                    btn_busca.config(image=nova_image_busca, fg='#7c7c7c')

                def volta_busca(e):
                    image_busca = Image.open('imagens\\lupa_solic.png')
                    resize_busca = image_busca.resize((25, 25))
                    nova_image_busca = ImageTk.PhotoImage(resize_busca)
                    btn_busca.photo = nova_image_busca
                    btn_busca.config(image=nova_image_busca, fg='#ffffff')

                image_busca = Image.open('imagens\\lupa_solic.png')
                resize_busca = image_busca.resize((25, 25))
                nova_image_busca = ImageTk.PhotoImage(resize_busca)
                btn_busca = Button(frame2, image=nova_image_busca, bg="#ffffff", fg='#FFFFFF', command=busca,
                                   borderwidth=0, relief=RIDGE, activebackground="#ffffff",
                                   activeforeground="#7c7c7c", cursor="hand2")
                btn_busca.photo = nova_image_busca
                btn_busca.grid(row=0, column=2)
                btn_busca.bind("<Enter>", muda_busca)
                btn_busca.bind("<Leave>", volta_busca)
                frame2.grid_columnconfigure(0, weight=1)
                frame2.grid_columnconfigure(3, weight=1)

                style = ttk.Style()
                # style.theme_use('default')
                style.configure('Treeview',
                                background='#ffffff',
                                rowheight=24,
                                fieldbackground='#ffffff',
                                font=fonte_padrao)
                style.configure("Treeview.Heading",
                                foreground='#000000',
                                background="#ffffff",
                                font=fonte_padrao)
                style.map('Treeview', background=[('selected', '#1d366c')])

                tree_principal = ttk.Treeview(frame3, selectmode='browse')
                vsb = ttk.Scrollbar(frame3, orient="vertical", command=tree_principal.yview)
                vsb.pack(side=RIGHT, fill='y')
                tree_principal.configure(yscrollcommand=vsb.set)
                # vsbx = ttk.Scrollbar(frame3, orient="horizontal", command=tree_principal.xview)
                # vsbx.pack(side=BOTTOM, fill='x')
                # tree_principal.configure(xscrollcommand=vsbx.set)
                tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
                tree_principal["columns"] = ("1")
                tree_principal['show'] = 'headings'
                tree_principal.column("1", width=300, anchor='c')
                tree_principal.heading("1", text="Nome do Solicitante")

                tree_principal.tag_configure('par', background='#e9e9e9')
                tree_principal.tag_configure('impar', background='#ffffff')
                tree_principal.bind("<Double-1>", duplo_adicionar_solicitante)
                frame3.grid_columnconfigure(0, weight=1)
                frame3.grid_columnconfigure(3, weight=1)

                bt2 = Button(frame4, text='Adicionar Solicitante', bg='#1d366c', fg='#FFFFFF',
                             activebackground='#1d366c',
                             activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                             command=adicionar_solicitante,
                             font=fonte_padrao, cursor="hand2")
                bt2.grid(row=0, column=2, pady=5, padx=5)

                frame4.grid_columnconfigure(0, weight=1)
                frame4.grid_columnconfigure(3, weight=1)

                conn = Connection("192.168.1.20", "gvdobrasil\\impressoras", "gv2K17ADM", auto_bind=True)
                conn.search('DC=gvdobrasil,DC=local',
                            "(&(objectClass=person)(objectClass=user)(sAMAccountType=805306368)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))",
                            SUBTREE, attributes=['sAMAccountName', 'displayName'])
                lista = []
                for i in conn.entries:
                    # result = '{0} {1}'.format(i.sAMAccountName.values, i.displayName.values)
                    conta_nome = format(i.sAMAccountName.values[0])
                    display_nome = format(i.displayName.values)
                    limpa = display_nome.translate({ord(i): None for i in "['],"})
                    lista.append((limpa))

                Label(frame6, text="Solicitante", font=fonte_padrao).grid(row=0, column=1, pady=(10, 0))
                ent_manual = Entry(frame6, width=20, font=fonte_padrao, justify='center')
                ent_manual.bind('<Return>', adicionar_solicitante_manual_bind)
                ent_manual.grid(row=1, column=1, padx=6)

                Label(frame6, text="E-mail", font=fonte_padrao).grid(row=0, column=2, pady=(10, 0))
                ent_email = Entry(frame6, width=32, font=fonte_padrao, justify='center')
                ent_email.bind('<Return>', adicionar_solicitante_manual_bind)
                ent_email.grid(row=1, column=2, padx=6)

                bt2 = Button(frame6, text='Adicionar Manualmente', bg='#1d366c', fg='#FFFFFF',
                             activebackground='#1d366c',
                             activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                             command=adicionar_solicitante_manual,
                             font=fonte_padrao, cursor="hand2")
                bt2.grid(row=2, column=1, pady=10, columnspan=2)

                frame6.grid_columnconfigure(0, weight=1)
                frame6.grid_columnconfigure(3, weight=1)

                Label(frame7, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1,
                                                                                           sticky="ew")
                frame7.grid_columnconfigure(0, weight=1)
                frame7.grid_columnconfigure(2, weight=1)
                '''root3.update()
                largura = frame0.winfo_width()
                altura = frame0.winfo_height()
                print(largura, altura)'''
                window_width = 496
                window_height = 518
                screen_width = root2.winfo_screenwidth()
                screen_height = root2.winfo_screenheight()
                x_cordinate = int((screen_width / 2) - (window_width / 2))
                y_cordinate = int((screen_height / 2) - (window_height / 2))
                root3.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
                root3.resizable(0, 0)
                root3.iconbitmap('imagens\\ico.ico')
                root3.title(titulo_todos)

            def enviar_int():
                global email_interacao
                email_interacao = 1
                txtinteracao.config(state='normal')
                if txtinteracao.get("1.0", 'end-1c') == "":
                    interacao = entryinteracao.get()
                    if interacao == "":
                        messagebox.showwarning('Campo vazio:', 'Atenção! Campo de Interação vazio.',
                                               parent=root2)
                        txtinteracao.config(state='disabled')
                    else:
                        hora = time.strftime('%H:%M:%S', time.localtime())
                        txtinteracao.insert(END,
                                            f'{usuariologado}: ({data} - {hora})\n{interacao}\n-----------------------------------------------------------------------')
                        txtinteracao.config(state='disabled')
                        entryinteracao.delete(0, END)
                else:
                    interacao = entryinteracao.get()
                    if interacao == "":
                        messagebox.showwarning('Campo vazio:', 'Atenção! Campo de Interação vazio.',
                                               parent=root2)
                        txtinteracao.config(state='disabled')
                    else:
                        hora = time.strftime('%H:%M:%S', time.localtime())
                        txtinteracao.insert(END,
                                            f'\n{usuariologado}: ({data} - {hora})\n{interacao}\n-----------------------------------------------------------------------')
                        txtinteracao.config(state='disabled')
                        entryinteracao.delete(0, END)

            def desfazer_int():
                entryinteracao.delete(0, END)

            def editar_chamado():
                root3 = Toplevel(root2)
                root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
                root3.unbind_class("Button", "<Key-space>")
                root3.focus_force()
                root3.grab_set()
                hora = time.strftime('%H:%M:%S', time.localtime())

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ ACESSOS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_acessos_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Liberação de acesso",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_acessos_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Sem permissão (Leitura ou Gravação)",
                        "Espaço insuficiente",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_acessos_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_acessos_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_acessos_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM ACESSOS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ HARDWARE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_hardware_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_hardware_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Máquina desligando",
                        "Máquina reiniciando",
                        "Não liga",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_hardware_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_hardware_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_hardware_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM HARDWARE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ TELEFONIA \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_telefonia_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Troca de Ramal",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_telefonia_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Não recebe ligação",
                        "Teclado com falha",
                        "Não liga",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_telefonia_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_telefonia_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_telefonia_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM TELEFONIA \\\\\\\\\\\\\\\\\\\\\\\\\\\

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ RADIO \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_radio_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        elif cliquesub.get() == "Envio de rádio para manutenção externa":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.delete('1.0', END)
                            txtdescr.insert(END,
                                            'Modelo:\n\nNº do Rádio:\n\nNº de Série:\n\nÁrea:\n\nResponsável:\n\nDefeito:\n\nCausa:')
                            txtdescr.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Compra",
                        "Configuração",
                        "Envio de rádio para manutenção externa",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_radio_problemas():
                    cliquetipo.set('')
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_radio_duvidas():
                    cliquetipo.set('')
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_radio_melhorias():
                    cliquetipo.set('')
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_radio_projetos():
                    cliquetipo.set('')
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM RADIO \\\\\\\\\\\\\\\\\\\\\\\\\\\

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ REDE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_rede_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_rede_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Sem acesso a rede (Dados)",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_rede_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_rede_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_rede_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM REDE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ INTERNET \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_internet_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Desbloqueio de Sites",
                        "Liberação WhatsAppWeb",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_internet_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Sem acesso",
                        "Lentidão",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_internet_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_internet_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_internet_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM INTERNET \\\\\\\\\\\\\\\\\\\\\\\\\\\
                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ IMPRESSORA \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_impressora_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Instalação",
                        "Troca de Toner",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_impressora_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Não está imprimindo",
                        "Enroscando papel",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_impressora_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_impressora_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_impressora_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM IMPRESSORA \\\\\\\\\\\\\\\\\\\\\\\\\\\
                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ E-MAIL \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_email_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Configuração",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_email_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Não envia e-mail",
                        "Não recebe e-mail",
                        "Pedindo senha",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_email_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_email_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_email_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM E-MAIL \\\\\\\\\\\\\\\\\\\\\\\\\\\
                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ SOFTWARES \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_softwares_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Instalação de Software",
                        "Opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_softwares_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Configuração",
                        "Travamento",
                        "Opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_softwares_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_softwares_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_softwares_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM SOFTWARES \\\\\\\\\\\\\\\\\\\\\\\\\\\
                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ WINDOWS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_windows_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Opções",
                        "Opções",
                        "Opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_windows_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Lentidão",
                        "Vírus",
                        "Opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_windows_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_windows_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_windows_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM WINDOWS \\\\\\\\\\\\\\\\\\\\\\\\\\\

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ PROTHEUS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                def opt_protheus_solicitacao():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Criação de usuário",
                        "Liberação de Módulo",
                        "Escolham algumas opções",
                        "Escolham algumas opções",
                        "Escolham algumas opções",
                        "Escolham algumas opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_protheus_problemas():
                    dropsub.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.config(state='disabled')

                    def clique(event):
                        if cliquesub.get() == "Outros assuntos..":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.focus_force()
                        else:
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.insert(0, cliquesub.get())
                            entrytitulo.config(state='disabled')
                            txtdescr.focus_force()

                    options = [
                        "Sem acesso",
                        "Trocar senha",
                        "Escolham algumas opções",
                        "Escolham algumas opções",
                        "Escolham algumas opções",
                        "Escolham algumas opções",
                        "Outros assuntos.."
                    ]
                    cliquesub = StringVar()
                    dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                    dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                             activeforeground="#FFFFFF",
                                             highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub_problemas.grid(row=1, column=3, sticky="w")

                def opt_protheus_duvidas():
                    # dropsub_solicitacao.grid_forget()
                    # dropsub_problemas.grid_forget()
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_protheus_melhorias():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                def opt_protheus_projetos():
                    entrytitulo.config(state='normal')
                    entrytitulo.delete(0, END)
                    entrytitulo.focus_force()
                    cliquesub = StringVar()
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.grid_forget()
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                   cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")

                # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM protheus \\\\\\\\\\\\\\\\\\\\\\\\\\\

                def dropselecaotipo(event):
                    if clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Solicitação":
                        opt_protheus_solicitacao()
                    elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Problemas":
                        opt_protheus_problemas()
                    elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Dúvidas":
                        opt_protheus_duvidas()
                    elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Melhorias":
                        opt_protheus_melhorias()
                    elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Projetos":
                        opt_protheus_projetos()
                    elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Solicitação":
                        opt_windows_solicitacao()
                    elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Problemas":
                        opt_windows_problemas()
                    elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Dúvidas":
                        opt_windows_duvidas()
                    elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Melhorias":
                        opt_windows_melhorias()
                    elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Projetos":
                        opt_windows_projetos()
                    elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Solicitação":
                        opt_softwares_solicitacao()
                    elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Problemas":
                        opt_softwares_problemas()
                    elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Dúvidas":
                        opt_softwares_duvidas()
                    elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Melhorias":
                        opt_softwares_melhorias()
                    elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Projetos":
                        opt_softwares_projetos()
                    elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Solicitação":
                        opt_email_solicitacao()
                    elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Problemas":
                        opt_email_problemas()
                    elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Dúvidas":
                        opt_email_duvidas()
                    elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Melhorias":
                        opt_email_melhorias()
                    elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Projetos":
                        opt_email_projetos()
                    elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Solicitação":
                        opt_impressora_solicitacao()
                    elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Problemas":
                        opt_impressora_problemas()
                    elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Dúvidas":
                        opt_impressora_duvidas()
                    elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Melhorias":
                        opt_impressora_melhorias()
                    elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Projetos":
                        opt_impressora_projetos()
                    elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Solicitação":
                        opt_internet_solicitacao()
                    elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Problemas":
                        opt_internet_problemas()
                    elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Dúvidas":
                        opt_internet_duvidas()
                    elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Melhorias":
                        opt_internet_melhorias()
                    elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Projetos":
                        opt_internet_projetos()
                    elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Solicitação":
                        opt_radio_solicitacao()
                    elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Problemas":
                        opt_radio_problemas()
                    elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Dúvidas":
                        opt_radio_duvidas()
                    elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Melhorias":
                        opt_radio_melhorias()
                    elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Projetos":
                        opt_radio_projetos()
                    elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Solicitação":
                        opt_rede_solicitacao()
                    elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Problemas":
                        opt_rede_problemas()
                    elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Dúvidas":
                        opt_rede_duvidas()
                    elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Melhorias":
                        opt_rede_melhorias()
                    elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Projetos":
                        opt_rede_projetos()
                    elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Solicitação":
                        opt_telefonia_solicitacao()
                    elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Problemas":
                        opt_telefonia_problemas()
                    elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Dúvidas":
                        opt_telefonia_duvidas()
                    elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Melhorias":
                        opt_telefonia_melhorias()
                    elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Projetos":
                        opt_telefonia_projetos()
                    elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Solicitação":
                        opt_hardware_solicitacao()
                    elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Problemas":
                        opt_hardware_problemas()
                    elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Dúvidas":
                        opt_hardware_duvidas()
                    elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Melhorias":
                        opt_hardware_melhorias()
                    elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Projetos":
                        opt_hardware_projetos()
                    elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Solicitação":
                        opt_acessos_solicitacao()
                    elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Problemas":
                        opt_acessos_problemas()
                    elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Dúvidas":
                        opt_acessos_duvidas()
                    elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Melhorias":
                        opt_acessos_melhorias()
                    elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Projetos":
                        opt_acessos_projetos()

                def dropselecao_ocorr(event):
                    if clique_ocorr.get() == "Protheus":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    elif clique_ocorr.get() == "Windows":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Softwares":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "E-mail":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Impressora":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Internet":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Rádio":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Rede":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Telefonia":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Hardware":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")
                    elif clique_ocorr.get() == "Acessos":
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquetipo.set('')
                        droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                        activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                        width=24)
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51,
                                       cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                def conversor(filename):
                    # Convert digital data to binary format
                    with open(filename, 'rb') as file:
                        blobData = file.read()
                    return blobData

                def anexo():
                    anexo = filedialog.askopenfilename(initialdir="os.path.expanduser(default_dir)",
                                                       title="Escolha um Arquivo",
                                                       filetypes=(
                                                           [("JPG", "*.jpg"), ("JPEG", "*.jpeg"), ("Bitmap", "*.bmp"),
                                                            ("PNG", "*.png"), ("Texto", "*.txt"), ("Word", "*.docx"),
                                                            ("Excel", "*.xlsx"), ("Outlook", "*.msg")]), parent=root2)
                    entryanexo.config(state=NORMAL)
                    entryanexo.insert(0, anexo)
                    entryanexo.config(state=DISABLED)
                    extensao = os.path.splitext(anexo)
                    global anexo_extensao
                    anexo_extensao = extensao[1]
                    global imgconvertida
                    imgconvertida = conversor(anexo)

                def confirmar_edicao():
                    if cliquetipo.get() == "" or entrytitulo.get() == "" or txtdescr.get("1.0", 'end-1c') == "":
                        messagebox.showwarning('+ Abrir Chamado: Erro',
                                               'Todos os campos com ( * ) devem ser preenchidos.',
                                               parent=root3)
                    else:
                        data_abertura = entrydataabertura.get()
                        ocorr = clique_ocorr.get().upper()
                        tipo = cliquetipo.get().upper()
                        titulo = entrytitulo.get().upper()
                        descricao_problema = txtdescr.get("1.0", 'end-1c')
                        nome_maquina = entrynomemaquina.get().upper()
                        ramal = entryramal.get().upper()
                        setor = entrysetor.get().upper()
                        resposta = messagebox.askyesno('Atenção:',
                                                       f'Tem certeza de que deseja editar o chamado nº{result[0]}?',
                                                       parent=root3)
                        if resposta == True:
                            cursor.execute(
                                "UPDATE helpdesk.dbo.chamados SET data_abertura = ? , ocorrencia = ?, tipo = ?,  titulo = ?, descricao_problema = ?, nome_maquina = ?, ramal = ?, setor = ? WHERE id_chamado = ?",
                                (data_abertura, ocorr, tipo, titulo, descricao_problema, nome_maquina, ramal, setor,
                                 n_chamado))
                            cursor.commit()
                            atualizar_lista_principal()
                            root3.destroy()
                            root2.destroy()

                def cancelar():
                    root2.destroy()

                def setup_entradas():
                    entryusuario.insert(0, result[1])
                    entryusuario.config(state='disabled')
                    entry_solicitante.insert(0, result[16])
                    entry_solicitante.config(state='disabled')
                    clique_ocorr.set(result[17])
                    cliquetipo.set(result[4])
                    clique_setor.set(result[10])
                    btnanexo.config(state='disabled')
                    entrytitulo.config(state='normal')
                    entrytitulo.insert(0, result[5])
                    txtdescr.insert(END, result[7])
                    entrynomemaquina.insert(0, result[8])
                    entryramal.insert(0, result[9])
                    entrysetor.insert(0, result[10])

                    if result[21] != None:
                        entryemail.config(state='normal')
                        entryemail.insert(0, result[21])
                        entryemail.config(state='disabled')

                frame0 = Frame(root3, bg='#ffffff')
                frame0.grid(row=0, column=0, stick='nsew')
                root2.grid_rowconfigure(0, weight=1)
                root2.grid_columnconfigure(0, weight=1)
                frame1 = Frame(frame0, bg="#1d366c")
                frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
                frame2 = Frame(frame0, bg='#ffffff')
                frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                frame3 = Frame(frame0, bg='#ffffff')
                frame3.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                frame4 = Frame(frame0, bg='#ffffff')
                frame4.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                frame5 = Frame(frame0, bg='#ffffff')
                frame5.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                frame6 = Frame(frame0, bg='#ffffff')
                frame6.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                frame7 = Frame(frame0, bg='#ffffff')
                frame7.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                frame8 = Frame(frame0, bg='#1d366c')
                frame8.pack(side=TOP, fill=X, expand=False, anchor='center')

                Label(frame1, image=nova_image_chamado, text=f" Editar Chamado: Nº {result[0]}",
                      compound="left",
                      bg='#1d366c',
                      fg='#FFFFFF',
                      font=fonte_titulos).grid(row=0, column=1)
                frame1.grid_columnconfigure(0, weight=1)
                frame1.grid_columnconfigure(2, weight=1)

                Label(frame2, text="Usuário Logado:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1,
                                                                                            sticky="w")
                entryusuario = Entry(frame2, font=fonte_padrao, justify='center')
                entryusuario.grid(row=1, column=1, sticky="w")

                Label(frame2, text="Solicitante: *", font=fonte_padrao, fg='#8B0000', bg='#ffffff').grid(row=0,
                                                                                                         column=2,
                                                                                                         sticky="w",
                                                                                                         padx=12)
                entry_solicitante = Entry(frame2, font=fonte_padrao, justify='center')
                entry_solicitante.grid(row=1, column=2, sticky="w", padx=12)
                entry_solicitante.focus_force()

                Label(frame2, text="Data de Abertura:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=3,
                                                                                              sticky="w",
                                                                                              padx=12)
                entrydataabertura = Entry(frame2, font=fonte_padrao, justify='center')
                entrydataabertura.grid(row=1, column=3, sticky="w", padx=12)
                entrydataabertura.insert(0, data)
                entrydataabertura.config(state='disabled')

                Label(frame2, text="E-mail:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=4,
                                                                                    sticky="w")
                entryemail = Entry(frame2, font=fonte_padrao, justify='center', width=36)
                entryemail.grid(row=1, column=4, sticky="w")
                entryemail.config(state='disabled')
                frame2.grid_columnconfigure(0, weight=1)
                frame2.grid_columnconfigure(5, weight=1)

                Label(frame3, text="Ocorrência: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0,
                                                                                                        column=1,
                                                                                                        sticky="w")
                clique_ocorr = StringVar()
                drop_ocorr = OptionMenu(frame3, clique_ocorr, "Protheus", "Windows", "Softwares", "E-mail",
                                        "Impressora", "Internet", "Rádio", "Rede", "Telefonia", "Hardware", "Acessos",
                                        command=dropselecao_ocorr)
                drop_ocorr.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                  activeforeground="#FFFFFF",
                                  highlightthickness=0, relief=RIDGE, width=24, cursor="hand2")
                drop_ocorr.grid(row=1, column=1, sticky="w")

                Label(frame3, text="Tipo: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0,
                                                                                                  column=2,
                                                                                                  sticky="w",
                                                                                                  padx=10)
                cliquetipo = StringVar()
                droptipo = OptionMenu(frame3, cliquetipo, "Solicitação", "Problemas", "Dúvidas", "Melhorias",
                                      "Projetos", command=dropselecaotipo)
                droptipo.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                activeforeground="#FFFFFF",
                                highlightthickness=0, relief=RIDGE, width=24, cursor="hand2")
                droptipo.grid(row=1, column=2, sticky="w", padx=10)
                Label(frame3, text="Título Predefinido: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(
                    row=0,
                    column=3,
                    sticky="w")
                cliquesub = StringVar()
                global dropsub
                dropsub = OptionMenu(frame3, cliquesub, "")
                dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                               activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                dropsub.grid(row=1, column=3, sticky="w")
                frame3.grid_columnconfigure(0, weight=1)
                frame3.grid_columnconfigure(4, weight=1)

                def muda_anexo(e):
                    image_anexo = Image.open('imagens\\anexo_over.png')
                    resize_anexo = image_anexo.resize((15, 20))
                    nova_image_anexo = ImageTk.PhotoImage(resize_anexo)
                    btnanexo.photo = nova_image_anexo
                    btnanexo.config(image=nova_image_anexo, fg='#7c7c7c')

                def volta_anexo(e):
                    image_anexo = Image.open('imagens\\anexo.png')
                    resize_anexo = image_anexo.resize((15, 20))
                    nova_image_anexo = ImageTk.PhotoImage(resize_anexo)
                    btnanexo.photo = nova_image_anexo
                    btnanexo.config(image=nova_image_anexo, fg='#1d366c')

                image_anexo = Image.open('imagens\\anexo.png')
                resize_anexo = image_anexo.resize((15, 20))
                nova_image_anexo = ImageTk.PhotoImage(resize_anexo)
                btnanexo = Button(frame4, image=nova_image_anexo, text=" Anexar arquivo.", compound="left",
                                  font=fonte_padrao, bg='#ffffff', fg='#1d366c', command=anexo,
                                  borderwidth=0, relief=RIDGE, activebackground="#ffffff",
                                  activeforeground="#7c7c7c", cursor="hand2")
                btnanexo.photo = nova_image_anexo
                btnanexo.grid(row=0, column=2, sticky="w", padx=(9, 0))
                btnanexo.bind("<Enter>", muda_anexo)
                btnanexo.bind("<Leave>", volta_anexo)
                lbltitulo = Label(frame4, text="Título: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000')
                lbltitulo.grid(row=0, column=1, sticky="w")
                entrytitulo = Entry(frame4, font=fonte_padrao, justify='center', width=66)
                entrytitulo.grid(row=1, column=1, sticky="ew", padx=(0, 9))
                entrytitulo.config(state=DISABLED)
                entryanexo = Entry(frame4, font=fonte_padrao, justify='center', width=36)
                entryanexo.grid(row=1, column=2, sticky="ew", padx=(9, 0))
                entryanexo.config(state='disabled')
                frame4.grid_columnconfigure(0, weight=1)
                frame4.grid_columnconfigure(3, weight=1)

                Label(frame5, text="Descrição do Problema: *", font=fonte_padrao, bg='#ffffff',
                      fg='#8B0000').grid(
                    row=0, column=1, sticky="w")
                txtdescr = scrolledtext.ScrolledText(frame5, width=103, height=10, font=fonte_padrao)
                txtdescr.grid(row=1, column=1)
                frame5.grid_columnconfigure(0, weight=1)
                frame5.grid_columnconfigure(2, weight=1)

                Label(frame6, text="Nome da Máquina:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(
                    row=0,
                    column=1,
                    sticky="w")
                entrynomemaquina = Entry(frame6, font=fonte_padrao, justify='center', width=25)
                entrynomemaquina.grid(row=1, column=1, sticky="w")

                Label(frame6, text="Ramal:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w",
                                                                                   padx=14)
                entryramal = Entry(frame6, font=fonte_padrao, justify='center', width=25)
                entryramal.grid(row=1, column=2, sticky="w", padx=14)

                Label(frame6, text="Setor: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0,
                                                                                                   column=3, sticky="w")
                OptionList = [
                    "Aciaria",
                    "Almoxarifado",
                    "Ambulatório",
                    "Balança",
                    "Comercial",
                    "Compras",
                    "Contabilidade",
                    "Custos",
                    "EHS",
                    "Elétrica",
                    "Engenharia",
                    "Faturamento",
                    "Financeiro",
                    "Fiscal",
                    "Lab Inspeção",
                    "Lab Mecânico",
                    "Lab Químico",
                    "Laminação",
                    "Logística",
                    "Oficina de Cilindros",
                    "Oficina Mecânica",
                    "Pátio de Sucata",
                    "PCP",
                    "Planta D´agua",
                    "Planta de Escória",
                    "Portaria",
                    "Qualidade",
                    "Refratários",
                    "RH",
                    "Subestação",
                    "TI",
                    "Utilidades"
                ]
                clique_setor = StringVar()
                drop_setor = OptionMenu(frame6, clique_setor, *OptionList)
                drop_setor.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                  activeforeground="#FFFFFF",
                                  highlightthickness=0, relief=RIDGE, width=25, cursor="hand2")
                drop_setor.grid(row=1, column=3, sticky="w")

                frame6.grid_columnconfigure(0, weight=1)
                frame6.grid_columnconfigure(4, weight=1)

                bt1 = Button(frame7, text='Confirmar', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                             activeforeground="#FFFFFF", highlightthickness=0, width=10, relief=RIDGE,
                             command=confirmar_edicao,
                             font=fonte_padrao, cursor="hand2")
                bt1.grid(row=0, column=1, padx=5)
                bt2 = Button(frame7, text='Cancelar', width=10, relief=RIDGE, command=cancelar,
                             font=fonte_padrao, cursor="hand2")
                bt2.grid(row=0, column=2, padx=5)
                frame7.grid_columnconfigure(0, weight=1)
                frame7.grid_columnconfigure(3, weight=1)

                Label(frame8, text=" ", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos).grid(row=0, column=1)
                frame8.grid_columnconfigure(0, weight=1)
                frame8.grid_columnconfigure(2, weight=1)
                setup_entradas()
                '''root2.update()
                largura = frame0.winfo_width()
                altura = frame0.winfo_height()
                print(largura, altura)'''
                window_width = 742
                window_height = 577
                screen_width = root2.winfo_screenwidth()
                screen_height = root2.winfo_screenheight()
                x_cordinate = int((screen_width / 2) - (window_width / 2))
                y_cordinate = int((screen_height / 2) - (window_height / 2))
                root3.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
                root3.resizable(0, 0)
                root3.configure(bg='#000000')
                root3.iconbitmap('imagens\\ico.ico')

            frame0 = Frame(root2, bg='#ffffff')
            frame0.grid(row=0, column=1, sticky=NSEW)
            frame_topo = Frame(frame0, bg='#1d366c')
            frame_topo.grid(row=0, column=1, sticky=NSEW, columnspan=2)
            frame_esquerda = LabelFrame(frame0, text=f"Informações do Chamado -- Aberto por: {result[1]}",
                                        font=fonte_padrao, bg='#ffffff')
            frame_esquerda.grid(row=1, column=1, sticky=NSEW, padx=10)
            frame_direita = Frame(frame0, bg='#ffffff')
            frame_direita.grid(row=1, column=2, sticky=NSEW, padx=10)
            frame_baixo = Frame(frame0, bg='#1d366c')
            frame_baixo.grid(row=2, column=1, sticky=NSEW, columnspan=2)
            # //// TOPO ////
            Label(frame_topo, image=nova_image_atendimento, text=f" Atendimento: Chamado Nº {n_chamado}",
                  compound="left", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos).grid(row=0, column=1)
            frame_topo.grid_columnconfigure(0, weight=1)
            frame_topo.grid_columnconfigure(2, weight=1)
            # //// ESQUERDA ////
            frame1 = Frame(frame_esquerda, bg='#ffffff')
            frame1.grid(row=0, column=1, sticky=EW, pady=(6, 0))
            frame2 = Frame(frame_esquerda, bg='#ffffff')
            frame2.grid(row=1, column=1, sticky=NSEW, pady=6)
            frame3 = Frame(frame_esquerda, bg='#ffffff')
            frame3.grid(row=2, column=1, sticky=NSEW, pady=6)
            frame4 = Frame(frame_esquerda, bg='#ffffff')
            frame4.grid(row=3, column=1, sticky=NSEW, pady=6)
            frame5 = Frame(frame_esquerda, bg='#ffffff')
            frame5.grid(row=4, column=1, sticky=NSEW, pady=6)
            frame6 = Frame(frame_esquerda, bg='#ffffff')
            frame6.grid(row=5, column=1, sticky=NSEW, pady=6)

            Label(frame1, text="Solicitante:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w",
                                                                                     padx=6)
            entrysolicitante = Entry(frame1, font=fonte_padrao, justify='center', bg='#ffffff')
            entrysolicitante.grid(row=1, column=1, sticky="w", padx=6)

            Label(frame1, text="Data de Abertura:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w",
                                                                                          padx=6)
            entrydataabertura = Entry(frame1, font=fonte_padrao, justify='center', bg='#ffffff')
            entrydataabertura.grid(row=1, column=2, sticky="w", padx=6)

            Label(frame1, text="E-mail:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=3, sticky="w", padx=6)
            entryemail = Entry(frame1, font=fonte_padrao, justify='center', bg='#ffffff', width=34)
            entryemail.grid(row=1, column=3, sticky="w", padx=6)

            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(4, weight=1)

            Label(frame2, text="Hora:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w", padx=6)
            entryhoraabertura = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=10)
            entryhoraabertura.grid(row=1, column=1, sticky="w", padx=6)

            Label(frame2, text="Nome da Máquina:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w",
                                                                                          padx=6)
            entrynomemaquina = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=20)
            entrynomemaquina.grid(row=1, column=2, sticky="w", padx=6)

            Label(frame2, text="Ramal:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=3, sticky="w", padx=6)
            entryramal = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=15)
            entryramal.grid(row=1, column=3, sticky="w", padx=6)

            Label(frame2, text="Setor:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=4, sticky="w", padx=6)
            entrysetor = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=27)
            entrysetor.grid(row=1, column=4, sticky="w", padx=6)

            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(5, weight=1)

            Label(frame3, text="Ocorrência:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w", padx=6)
            entryocorrencia = Entry(frame3, font=fonte_padrao, justify='center', bg='#ffffff', width=26)
            entryocorrencia.grid(row=1, column=1, sticky="ew", padx=6)

            Label(frame3, text="Tipo:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w", padx=6)
            entrytipo = Entry(frame3, font=fonte_padrao, justify='center', bg='#ffffff', width=50)
            entrytipo.grid(row=1, column=2, columnspan=2, sticky="w", padx=6)

            frame3.grid_columnconfigure(0, weight=1)
            frame3.grid_columnconfigure(3, weight=1)

            lbltitulo = Label(frame4, text="Título:*", font=fonte_padrao, bg='#ffffff')
            lbltitulo.grid(row=0, column=1, sticky="w")
            entrytitulo = Entry(frame4, font=fonte_padrao, justify='center', width=79)
            entrytitulo.grid(row=1, column=1, sticky="w", columnspan=3)

            frame4.grid_columnconfigure(0, weight=1)
            frame4.grid_columnconfigure(3, weight=1)

            Label(frame5, text="Descrição do Problema:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1,
                                                                                                sticky="w")
            txtdescr_atendimento = scrolledtext.ScrolledText(frame5, width=77, height=16, font=fonte_padrao,
                                                             bg='#ffffff', wrap=WORD)
            txtdescr_atendimento.grid(row=1, column=1, sticky="ew")

            frame5.grid_columnconfigure(0, weight=1)
            frame5.grid_columnconfigure(2, weight=1)

            def muda_anexo(e):
                image_abriranexo = Image.open('imagens\\anexo_over.png')
                resize_abriranexo = image_abriranexo.resize((15, 20))
                nova_image_abriranexo = ImageTk.PhotoImage(resize_abriranexo)
                btabrir_anexo.photo = nova_image_abriranexo
                btabrir_anexo.config(image=nova_image_abriranexo, fg='#7c7c7c')

            def volta_anexo(e):
                # image_abriranexo = Image.open('imagens\\anexo.png')
                resize_abriranexo = image_abriranexo.resize((15, 20))
                nova_image_abriranexo = ImageTk.PhotoImage(resize_abriranexo)
                btabrir_anexo.photo = nova_image_abriranexo
                btabrir_anexo.config(image=nova_image_abriranexo, fg='#1d366c')

            image_abriranexo = Image.open('imagens\\anexo.png')
            resize_abriranexo = image_abriranexo.resize((15, 20))
            nova_image_abriranexo = ImageTk.PhotoImage(resize_abriranexo)
            btabrir_anexo = Button(frame6, image=nova_image_abriranexo, text=" Abrir Anexo.", compound="left",
                                   font=fonte_padrao, bg='#ffffff', fg='#1d366c', command=abrir_anexo,
                                   borderwidth=0, relief=RIDGE, activebackground="#ffffff",
                                   activeforeground="#7c7c7c", cursor="hand2")
            btabrir_anexo.photo = nova_image_abriranexo
            btabrir_anexo.grid(row=0, column=0, padx=(6, 0))
            btabrir_anexo.bind("<Enter>", muda_anexo)
            btabrir_anexo.bind("<Leave>", volta_anexo)

            bt_editar_chamado = Button(frame6, text='Editar', width=10, relief=RIDGE, command=editar_chamado,
                                       font=fonte_padrao, bg="#1d366c", fg="#ffffff",
                                       activebackground="#1d366c",
                                       activeforeground="#FFFFFF", state=NORMAL, cursor="hand2")
            bt_editar_chamado.grid(row=0, column=3, padx=(0, 6))

            frame6.grid_columnconfigure(1, weight=1)
            frame6.grid_columnconfigure(2, weight=1)
            # //// DIREITA ////
            frame1 = Frame(frame_direita, bg='#ffffff')
            frame1.grid(row=0, column=1, sticky=EW)
            frame2 = Frame(frame_direita, bg='#ffffff')
            frame2.grid(row=1, column=1, sticky=NSEW, pady=6)
            frame3 = LabelFrame(frame_direita, text="Interação do Chamado (Usuário\Analistas):",
                                font=fonte_padrao, bg='#ffffff')
            frame3.grid(row=2, column=1, sticky=NSEW, pady=6)
            frame4 = Frame(frame_direita, bg='#ffffff')
            frame4.grid(row=3, column=1, sticky=NSEW)
            frame5 = Frame(frame_direita, bg='#ffffff')
            frame5.grid(row=4, column=1, sticky=NSEW)

            Label(frame1, text="Nome do Analista:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1,
                                                                                          sticky="w")
            entryanalista2 = Entry(frame1, font=fonte_padrao, justify='center', width=70, bg='#ffffff')
            entryanalista2.grid(row=1, column=1, sticky="ew")
            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(2, weight=1)

            def cliquestatus(event):
                if clique_status.get() == "Encerrado" or clique_status.get() == "Cancelado":
                    entrydataencerramento.config(state='normal')
                    entrydataencerramento.delete(0, END)
                    entrydataencerramento.insert(0, data)
                    entrydataencerramento.config(state='disabled')
                elif clique_status.get() == "Em andamento" or clique_status.get() == "Aguardando Usuário" or clique_status.get() == "Chamado TOTVS":
                    entrydataencerramento.config(state='normal')
                    entrydataencerramento.delete(0, END)
                    entrydataencerramento.config(state='disabled')

            Label(frame2, text="Prioridade do Chamado:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1)
            options = ["0 - Baixa", "1 - Média", "2 - Alta", "3 - Urgente"]
            clique_prioridade = StringVar()

            drop_status = OptionMenu(frame2, clique_prioridade, *options)
            drop_status.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF",
                               highlightthickness=0, relief=RIDGE, width=20, cursor="hand2")
            drop_status.grid(row=1, column=1)

            Label(frame2, text="Encaminhar Atendimento", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2)

            def drop_redireciona(event):
                redirec = clique_redirec.get()
                entryanalista2.config(state='normal')
                entryanalista2.delete(0, END)
                entryanalista2.insert(0, redirec)
                entryanalista2.config(state='disabled')

            r = cursor.execute("SELECT * FROM dbo.analista")
            lista_analistas = []
            for row in r:
                lista_analistas.append(row[2])

            clique_redirec = StringVar()
            drop_redirec = OptionMenu(frame2, clique_redirec, *lista_analistas, command=drop_redireciona)
            drop_redirec.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                activeforeground="#FFFFFF",
                                highlightthickness=0, relief=RIDGE, width=20, cursor="hand2")
            drop_redirec.grid(row=1, column=2, padx=10)

            image_solicitante = Image.open('imagens\\seta.png')
            resize_solicitante = image_solicitante.resize((20, 20))
            nova_image_solicitante = ImageTk.PhotoImage(resize_solicitante)
            btnsolicitante = Button(frame2, image=nova_image_solicitante, text="Enviar Chamado: ", compound="right",
                                    font=fonte_padrao, bg='#ffffff', fg='#1d366c', command=encaminhar,
                                    borderwidth=0, relief=RIDGE, activebackground="#ffffff",
                                    activeforeground="#7c7c7c", cursor="hand2")
            btnsolicitante.photo = nova_image_solicitante
            btnsolicitante.grid(row=0, column=3)

            entry_enviar_chamado = Entry(frame2, font=fonte_padrao, justify='center', width=20, bg='#ffffff')
            entry_enviar_chamado.grid(row=1, column=3)

            Label(frame2, text="Status:", font=fonte_padrao, bg='#ffffff').grid(row=2, column=1, sticky="w")
            options = ["Em andamento", "Aguardando Usuário", "Chamado TOTVS", "Cancelado", "Encerrado"]
            clique_status = StringVar()

            drop_status = OptionMenu(frame2, clique_status, *options, command=cliquestatus)
            drop_status.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                               activeforeground="#FFFFFF",
                               highlightthickness=0, relief=RIDGE, width=20, cursor="hand2")
            drop_status.grid(row=3, column=1)

            Label(frame2, text="Data do Atendimento:", font=fonte_padrao, bg='#ffffff').grid(row=2, column=2)
            entrydataatendimento2 = Entry(frame2, font=fonte_padrao, justify='center', width=22, bg='#ffffff')
            entrydataatendimento2.grid(row=3, column=2)
            entrydataatendimento2.config(state='disabled')

            Label(frame2, text="Data de Encerramento:", font=fonte_padrao, bg='#ffffff').grid(row=2, column=3,
                                                                                              sticky="w")

            entrydataencerramento = Entry(frame2, font=fonte_padrao, justify='center', width=22, bg='#ffffff')
            entrydataencerramento.grid(row=3, column=3, columnspan=2, sticky="we")
            entrydataencerramento.config(state='disabled')

            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(5, weight=1)

            entryinteracao = Entry(frame3, font=fonte_padrao, justify='center', width=46, bg='#ffffff')
            entryinteracao.grid(row=0, column=1, padx=(2, 0))

            btn_envia_interacao = Button(frame3, text='Enviar', width=10, relief=RIDGE, font=fonte_padrao,
                                         command=enviar_int, bg="#1d366c", fg="#ffffff",
                                         activebackground="#1d366c", activeforeground="#FFFFFF", state=NORMAL,
                                         cursor="hand2")
            btn_envia_interacao.grid(row=0, column=2, padx=(5, 0))
            btn_cancela_interacao = Button(frame3, text='Desfazer', width=10, font=fonte_padrao,
                                           command=desfazer_int, cursor="hand2")
            btn_cancela_interacao.grid(row=0, column=3, padx=(5, 2))

            Label(frame3, text="Histórico:", font=fonte_padrao, bg='#ffffff').grid(row=1, column=1, sticky="w")
            txtinteracao = scrolledtext.ScrolledText(frame3, font=fonte_padrao, width=68, height=10,
                                                     bg='#ffffff', wrap=WORD)
            txtinteracao.grid(row=2, column=1, columnspan=3, sticky="ew", padx=2, pady=(0, 2))
            txtinteracao.config(state='disabled', fg='#696969')

            frame3.grid_columnconfigure(0, weight=1)
            frame3.grid_columnconfigure(4, weight=1)

            Label(frame4, text="Solução:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w")
            txtsolucao = scrolledtext.ScrolledText(frame4, width=71, height=5, font=fonte_padrao, bg='#ffffff',
                                                   wrap=WORD)
            txtsolucao.grid(row=1, column=1, sticky="ew", columnspan=2)
            txtsolucao.focus_force()
            frame4.grid_columnconfigure(0, weight=1)
            frame4.grid_columnconfigure(3, weight=1)

            btcancelar_atendimento = Button(frame5, text='Salvar', width=10, relief=RIDGE, command=salvar,
                                            font=fonte_padrao, bg="#1d366c", fg="#ffffff",
                                            activebackground="#1d366c", activeforeground="#FFFFFF",
                                            state=NORMAL, cursor="hand2")
            btcancelar_atendimento.grid(row=0, column=1, padx=5, pady=6)
            btconfirma_atencimento = Button(frame5, text='Cancelar', command=cancelar, width=10,
                                            font=fonte_padrao, cursor="hand2")
            btconfirma_atencimento.grid(row=0, column=2, padx=5, pady=6)

            frame5.grid_columnconfigure(0, weight=1)
            frame5.grid_columnconfigure(3, weight=1)

            # //// BAIXO ////
            Label(frame_baixo, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos).grid(row=0, column=1)
            frame_baixo.grid_columnconfigure(0, weight=1)
            frame_baixo.grid_columnconfigure(2, weight=1)
            setup_entradas()
            '''root2.update()
            largura = frame0.winfo_width()
            altura = frame0.winfo_height()
            print(largura, altura)'''
            window_width = 1136
            window_height = 657
            screen_width = root2.winfo_screenwidth()
            screen_height = root2.winfo_screenheight()
            x_cordinate = int((screen_width / 2) - (window_width / 2))
            y_cordinate = int((screen_height / 2) - (window_height / 2))
            root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
            root2.resizable(0, 0)
            root2.configure(bg='#ffffff')
            root2.iconbitmap('imagens\\ico.ico')
            root2.grid_columnconfigure(0, weight=1)
            root2.grid_columnconfigure(2, weight=1)
        ##controle da variavel de e-mail.. caso seja 1 o sistema enviara um email na interação
        global email_interacao
        email_interacao = 0
        global troca_analista
        troca_analista = 0
        chamado_select = tree_principal.focus()
        if chamado_select == "":
            messagebox.showwarning('Atendimento:', 'Selecione um chamado na lista!', parent=root)
        else:
            n_chamado = tree_principal.item(chamado_select, "values")[0]
            r = cursor.execute("SELECT * FROM dbo.chamados WHERE id_chamado=?", (n_chamado,))
            result = r.fetchone()
            if result[11] == 'Encerrado' or result[11] == 'Cancelado':
                messagebox.showwarning('Atenção:',
                                       'Este chamado já foi finalizado. Acesse "Visualizar Atendimento" para eventuais consultas.', parent=root)
            elif result[13] != usuariologado and result[13] != None:
                resposta = messagebox.askyesno('Atendimento:', f'Este chamado já está sendo atendido pelo(a) analista:\n{result[13]}.\n\nTem certeza que deseja passar a atender este chamado?', parent=root)
                if resposta == True:
                    troca_analista = 1
                    layout()
            else:
                layout()
    # /////////////////////////////FIM ATENDIMENTO/////////////////////////////

    # /////////////////////////////VISUALIZAR/////////////////////////////
    def visualizar_chamado_bind(event):
        visualizar_chamado()
    def visualizar_chamado():
        chamado_select = tree_principal.focus()
        if chamado_select == "":
            messagebox.showwarning('Atendimento:', 'Selecione um chamado na lista!', parent=root)
        else:
            n_chamado = tree_principal.item(chamado_select, "values")[0]
            r = cursor.execute("SELECT * FROM dbo.chamados WHERE id_chamado=?", (n_chamado,))
            result = r.fetchone()
            root2 = Toplevel(root)
            root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
            root2.unbind_class("Button", "<Key-space>")
            root2.focus_force()
            root2.grab_set()
            def setup_entradas():
                if nivel_acesso == 1: #Analista
                    bt_editar_chamado.config(state='disabled')
                    bt_excluir_chamado.config(state='disabled')
                    if result[16] != None:
                        entrysolicitante.insert(0, result[16])
                        entrysolicitante.config(state='disabled')
                    entrydataabertura.insert(0, result[2])
                    entrydataabertura.config(state='disabled')
                    if result[3] == None:
                        entryhoraabertura.insert(0, "")
                        entryhoraabertura.config(state='disabled')
                    else:
                        entryhoraabertura.insert(0, result[3])
                        entryhoraabertura.config(state='disabled')
                    if result[21] != None:
                        entryemail.insert(0, result[21])
                        entryemail.config(state='disabled')
                    entrynomemaquina.insert(0, result[8])
                    entrynomemaquina.config(state='disabled')
                    entryramal.insert(0, result[9])
                    entryramal.config(state='disabled')
                    entrysetor.insert(0, result[10])
                    entrysetor.config(state='disabled')
                    if result[17] != None:
                        entryocorrencia.insert(0, result[17])
                        entryocorrencia.config(state='disabled')
                    entrytitulo.insert(0, result[5])
                    entrytitulo.config(state='disabled')
                    entrytipo.insert(0, result[4])
                    entrytipo.config(state='disabled')
                    if result[13] != None:
                        entryanalista2.insert(0, result[13])
                        entryanalista2.config(state='disabled')
                    else:
                        entryanalista2.config(state='disabled')
                    txtdescr_atendimento.insert(END, result[7])
                    txtdescr_atendimento.config(state='disabled', fg='#696969')
                    btn_calendario_atendimento.config(state='disabled')
                    btn_calendario_encerramento.config(state='disabled')
                    if result[14] != None:
                        txtsolucao.config(state='normal')
                        txtsolucao.insert(END, result[14])
                        txtsolucao.config(state='disabled', fg='#696969')
                    if result[18] != None:
                        clique_prioridade.set(result[18])
                        drop_prioridade.config(state='disabled')
                    else:
                        drop_prioridade.config(state='disabled')

                    if result[6] == None and result[22] == None:
                        btabrir_anexo.config(state='disabled')

                    if result[12] != None:
                        entrydataatendimento2.config(state='normal')
                        entrydataatendimento2.insert(0, result[12])
                        entrydataatendimento2.config(state='disabled')
                    clique_status.set(result[11])
                    drop_status.config(state='disabled')
                    if result[15] != None:
                        entrydataencerramento.config(state='normal')
                        entrydataencerramento.insert(0, result[15])
                        entrydataencerramento.config(state='disabled')
                    clique_status.set(result[11])
                    drop_status.config(state='disabled')
                    entryinteracao.config(state='disabled')
                    btn_envia_interacao.config(state='disabled')
                    btn_cancela_interacao.config(state='disabled')
                    btconfirma_atencimento.config(state='disabled')
                    btcancelar_atendimento.config(state='disabled')
                    if result[19] != None:
                        txtinteracao.config(state='normal')
                        txtinteracao.insert(END, result[19])
                        txtinteracao.config(state='disabled')
                    upper = usuariologado.upper()
                    if result[11] == "Aberto" and result[1] == upper:
                        bt_editar_chamado.config(state='normal')
                        bt_excluir_chamado.config(state='normal')
                else: #Usuários
                    if result[16] != None:
                        entrysolicitante.insert(0, result[16])
                        entrysolicitante.config(state='disabled')
                    entrydataabertura.insert(0, result[2])
                    entrydataabertura.config(state='disabled')

                    if result[3] == None:
                        entryhoraabertura.insert(0, "")
                        entryhoraabertura.config(state='disabled')
                    else:
                        entryhoraabertura.insert(0, result[3])
                        entryhoraabertura.config(state='disabled')

                    if result[21] != None:
                        entryemail.insert(0, result[21])
                        entryemail.config(state='disabled')
                    entrynomemaquina.insert(0, result[8])
                    entrynomemaquina.config(state='disabled')
                    entryramal.insert(0, result[9])
                    entryramal.config(state='disabled')
                    entrysetor.insert(0, result[10])
                    entrysetor.config(state='disabled')
                    if result[17] != None:
                        entryocorrencia.insert(0, result[17])
                        entryocorrencia.config(state='disabled')

                    entrytitulo.insert(0, result[5])
                    entrytitulo.config(state='disabled')
                    entrytipo.insert(0, result[4])
                    entrytipo.config(state='disabled')

                    if result[13] != None:
                        entryanalista2.insert(0, result[13])
                        entryanalista2.config(state='disabled')
                    else:
                        entryanalista2.config(state='disabled')

                    txtdescr_atendimento.insert(END, result[7])
                    txtdescr_atendimento.config(fg='#696969')
                    txtdescr_atendimento.config(state='disabled')
                    btn_calendario_atendimento.config(state='disabled')
                    btn_calendario_encerramento.config(state='disabled')
                    if result[14] != None:
                        txtsolucao.config(state='normal')
                        txtsolucao.insert(END, result[14])
                        txtsolucao.config(state='disabled', fg='#696969')

                    if result[18] != None:
                        clique_prioridade.set(result[18])
                        drop_prioridade.config(state='disabled')
                    else:
                        drop_prioridade.config(state='disabled')

                    if result[6] == None and result[22] == None:
                        btabrir_anexo.config(state='disabled')

                    if result[12] != None:
                        entrydataatendimento2.config(state='normal')
                        entrydataatendimento2.insert(0, result[12])
                        entrydataatendimento2.config(state='disabled')

                    clique_status.set(result[11])
                    drop_status.config(state='disabled')
                    if result[11] != "Aberto":
                        bt_editar_chamado.config(state='disabled')
                        bt_excluir_chamado.config(state='disabled')
                    if result[11] == "Aberto":
                        entryinteracao.config(state='disabled')
                        btn_envia_interacao.config(state='disabled')
                        btn_cancela_interacao.config(state='disabled')
                    if result[11] == "Encerrado":
                        entrydataencerramento.config(state='normal')
                        entrydataencerramento.insert(0, result[15])
                        entrydataencerramento.config(state='disabled')
                        entryinteracao.config(state='disabled')
                        btn_envia_interacao.config(state='disabled')
                        btn_cancela_interacao.config(state='disabled')

                    if result[19] != None:
                        txtinteracao.config(state='normal')
                        txtinteracao.insert(END, result[19])
                        txtinteracao.config(state='disabled')

            def abrir_anexo():
                if result[6] != None:
                    def conversorarquivo(data, filename):
                        with open(filename, 'wb') as file:
                            file.write(data)

                    nome_arquivo = result[6]
                    abreviacao_extensao = result[20].upper()
                    nome_extensao = result[20]
                    caminho = filedialog.asksaveasfilename(defaultextension=".*",
                                                           initialfile='Anexo_Chamado (' + str(result[0]) + ')',
                                                           initialdir='os.path.expanduser(default_dir)',
                                                           title='Salvar Anexo..',
                                                           filetypes=[
                                                               (str(abreviacao_extensao), "*" + nome_extensao)], parent=root2)
                    conversorarquivo(nome_arquivo, caminho)

                else:
                    subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
                    subprocess.call(r'net use \\192.168.1.19 /user:impressoras gv2K17ADM', shell=True)

                    def conversorarquivo(origem, destino):
                        try:
                            shutil.copy(origem, destino)
                        except:
                            messagebox.showerror('Erro:',
                                                 'Arquivo de destino não encontrado ou Download cancelado.',
                                                 parent=root2)
                            return False
                        messagebox.showinfo('Download:', 'Download concluído com sucesso', parent=root2)

                    cursor.execute("SELECT * FROM dbo.chamados WHERE nome_anexo=? AND id_chamado=?",
                                   (result[22], result[0],))
                    busca = cursor.fetchone()
                    nome_pasta = result[23]
                    pasta = r'\\192.168.1.19/helpdesk/anexos/' + nome_pasta
                    caminhocompleto = os.path.join(pasta, busca[22])
                    nome = busca[22]
                    destino = filedialog.asksaveasfilename(
                        initialfile=str(nome),
                        initialdir='os.path.expanduser(default_dir)',
                        title='Salvar Anexo..',
                        filetypes=[("Todos os arquivos", "*")], parent=root2)
                    conversorarquivo(caminhocompleto, destino)

                    subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
            def salvar():
                #drop_status.config(state='normal')
                status = clique_status.get()
                interacao = txtinteracao.get("1.0", 'end-1c')
                
                cursor.execute("UPDATE helpdesk.dbo.chamados SET interacao = ? , status = ? WHERE id_chamado = ?",(interacao, status, n_chamado))
                cursor.commit()
                messagebox.showinfo('Atendimento:', 'Alteração realizada com sucesso!', parent=root2)
                atualizar_lista_principal()
                if email_interacao == 1:
                    cursor.execute("SELECT * FROM analista WHERE nome_analista = ?",(result[13]))
                    resultado = cursor.fetchone()
                    email_analista = resultado[4]
                    interacao_html = interacao.replace("\n", "<br>")

                    sender_email = "naoresponder@gruposimec.com.br"
                    receiver_email = email_analista

                    password = "Qu@@258147"

                    message = MIMEMultipart("alternative")
                    message["Subject"] = "HelpDesk - Interação do Chamado nº" + str(
                        result[0]) + " - (Não responder)"
                    message["From"] = sender_email
                    message["To"] = receiver_email

                    # Create the plain-text and HTML version of your message
                    text = """\
                    O colaborador  """ + str(result[16]) + """ interagiu ao seu chamado de nº """ + str(result[0]) + """.

                    Acesse o sistema o mais breve possível e responda a interação."""

                    html = """\
                    <html>
                        <body>
                        <center>	
                    <font size="2" face="Arial" >
                    <table width=100% border=0>
                        <tr style="background-color:#01336e">
                        <td align=center><p style= "font-family:Arial; font-size:50px; color:white"><b>HelpDesk</b></p></td>
                        </tr>
                        <tr>
                        <td align=center>O colaborador <b>""" + str(
                        result[16]) + """</b> interagiu ao seu chamado de nº <b>""" + str(result[0]) + """</b> e aguarda sua resposta. </td>
                        </tr>
                        <tr>
                            <td>Histórico da interação<br><br>"""+str(interacao_html)+"""</td>
                        </tr>
                        <tr>
                        <td  align=center><br>Acesse o sistema o mais breve possível e responda a interação.</td>
                        </tr>
                    </table>
                    </font>
                    </center>	

                        </body>
                    </html>

                    """

                    # Turn these into plain/html MIMEText objects
                    part1 = MIMEText(text, "plain")
                    part2 = MIMEText(html, "html")

                    # Add HTML/plain-text parts to MIMEMultipart message
                    # The email client will try to render the last part first
                    message.attach(part1)
                    message.attach(part2)

                    # Create secure connection with server and send email
                    context = ssl.create_default_context()
                    with smtplib.SMTP_SSL("smtps.uhserver.com", 465, context=context) as server:
                        server.login(sender_email, password)
                        server.sendmail(
                            sender_email, receiver_email, message.as_string()
                        )

                root2.destroy()
            def cancelar():
                root2.destroy()
            def enviar_int():
                global email_interacao
                email_interacao = 1
                txtinteracao.config(state='normal')
                if txtinteracao.get("1.0", 'end-1c') == "":
                    interacao = entryinteracao.get()
                    if interacao == "":
                        messagebox.showwarning('Campo vazio:', 'Atenção! Campo de Interação vazio.', parent=root2)
                        txtinteracao.config(state='disabled')
                    else:
                        hora = time.strftime('%H:%M:%S', time.localtime())
                        txtinteracao.insert(END,
                                            f'{usuariologado}: ({data} - {hora})\n{interacao}\n-----------------------------------------------------------------------')
                        txtinteracao.config(state='disabled')
                        entryinteracao.delete(0, END)
                        if nivel_acesso == 0:
                            drop_status.config(state='normal')
                            clique_status.set('Aguardando Analista')
                            drop_status.config(state='disabled')
                else:
                    interacao = entryinteracao.get()
                    if interacao == "":
                        messagebox.showwarning('Campo vazio:', 'Atenção! Campo de Interação vazio.', parent=root2)
                        txtinteracao.config(state='disabled')
                    else:
                        hora = time.strftime('%H:%M:%S', time.localtime())
                        txtinteracao.insert(END,
                                            f'\n{usuariologado}: ({data} - {hora})\n{interacao}\n-----------------------------------------------------------------------')
                        txtinteracao.config(state='disabled')
                        entryinteracao.delete(0, END)
                        if nivel_acesso == 0:
                            drop_status.config(state='normal')
                            clique_status.set('Aguardando Analista')
                            drop_status.config(state='disabled')


            def desfazer_int():
                entryinteracao.delete(0, END)
                txtinteracao.config(state='normal')
                txtinteracao.delete('1.0', END)
                txtinteracao.insert(END, result[19])
                txtinteracao.config(state='disabled')
            def excluir_chamado():
                resposta = messagebox.askyesno('Atenção:', f'Tem certeza de que deseja excluir o chamado nº{result[0]}?', parent=root2)
                if resposta == True:
                    cursor.execute("DELETE FROM helpdesk.dbo.chamados WHERE id_chamado = ?", (result[0]))
                    cursor.commit()
                    messagebox.showinfo('Atenção:', f'Chamado nº{result[0]} excluído com sucesso!', parent=root2)
                    atualizar_lista_principal()
                    root2.destroy()
            def editar_chamado():
                    root3 = Toplevel(root2)
                    root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
                    root3.unbind_class("Button", "<Key-space>")
                    root3.focus_force()
                    root3.grab_set()
                    hora = time.strftime('%H:%M:%S', time.localtime())

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ ACESSOS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_acessos_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Liberação de acesso",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_acessos_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Sem permissão (Leitura ou Gravação)",
                            "Espaço insuficiente",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_acessos_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_acessos_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_acessos_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM ACESSOS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ HARDWARE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_hardware_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_hardware_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Máquina desligando",
                            "Máquina reiniciando",
                            "Não liga",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_hardware_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_hardware_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_hardware_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM HARDWARE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ TELEFONIA \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_telefonia_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Troca de Ramal",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_telefonia_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Não recebe ligação",
                            "Teclado com falha",
                            "Não liga",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_telefonia_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_telefonia_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_telefonia_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM TELEFONIA \\\\\\\\\\\\\\\\\\\\\\\\\\\

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ RADIO \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_radio_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            elif cliquesub.get() == "Envio de rádio para manutenção externa":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.delete('1.0', END)
                                txtdescr.insert(END,
                                                'Modelo:\n\nNº do Rádio:\n\nNº de Série:\n\nÁrea:\n\nResponsável:\n\nDefeito:\n\nCausa:')
                                txtdescr.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Compra",
                            "Configuração",
                            "Envio de rádio para manutenção externa",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_radio_problemas():
                        cliquetipo.set('')
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_radio_duvidas():
                        cliquetipo.set('')
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_radio_melhorias():
                        cliquetipo.set('')
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_radio_projetos():
                        cliquetipo.set('')
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM RADIO \\\\\\\\\\\\\\\\\\\\\\\\\\\

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ REDE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_rede_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_rede_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Sem acesso a rede (Dados)",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_rede_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_rede_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_rede_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM REDE \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ INTERNET \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_internet_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Desbloqueio de Sites",
                            "Liberação WhatsAppWeb",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_internet_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Sem acesso",
                            "Lentidão",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_internet_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_internet_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_internet_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM INTERNET \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ IMPRESSORA \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_impressora_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Instalação",
                            "Troca de Toner",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_impressora_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Não está imprimindo",
                            "Enroscando papel",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_impressora_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_impressora_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_impressora_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM IMPRESSORA \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ E-MAIL \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_email_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Configuração",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_email_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Não envia e-mail",
                            "Não recebe e-mail",
                            "Pedindo senha",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_email_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_email_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_email_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM E-MAIL \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ SOFTWARES \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_softwares_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Instalação de Software",
                            "Opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_softwares_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Configuração",
                            "Travamento",
                            "Opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_softwares_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_softwares_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_softwares_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM SOFTWARES \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ WINDOWS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_windows_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Opções",
                            "Opções",
                            "Opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_windows_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Lentidão",
                            "Vírus",
                            "Opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_windows_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_windows_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_windows_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM WINDOWS \\\\\\\\\\\\\\\\\\\\\\\\\\\

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ PROTHEUS \\\\\\\\\\\\\\\\\\\\\\\\\\\
                    def opt_protheus_solicitacao():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Criação de usuário",
                            "Liberação de Módulo",
                            "Escolham algumas opções",
                            "Escolham algumas opções",
                            "Escolham algumas opções",
                            "Escolham algumas opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_protheus_problemas():
                        dropsub.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.config(state='disabled')

                        def clique(event):
                            if cliquesub.get() == "Outros assuntos..":
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.focus_force()
                            else:
                                entrytitulo.config(state='normal')
                                entrytitulo.delete(0, END)
                                entrytitulo.insert(0, cliquesub.get())
                                entrytitulo.config(state='disabled')
                                txtdescr.focus_force()

                        options = [
                            "Sem acesso",
                            "Trocar senha",
                            "Escolham algumas opções",
                            "Escolham algumas opções",
                            "Escolham algumas opções",
                            "Escolham algumas opções",
                            "Outros assuntos.."
                        ]
                        cliquesub = StringVar()
                        dropsub_problemas = OptionMenu(frame3, cliquesub, *options, command=clique)
                        dropsub_problemas.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                                 activeforeground="#FFFFFF",
                                                 highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub_problemas.grid(row=1, column=3, sticky="w")

                    def opt_protheus_duvidas():
                        # dropsub_solicitacao.grid_forget()
                        # dropsub_problemas.grid_forget()
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_protheus_melhorias():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    def opt_protheus_projetos():
                        entrytitulo.config(state='normal')
                        entrytitulo.delete(0, END)
                        entrytitulo.focus_force()
                        cliquesub = StringVar()
                        dropsub = OptionMenu(frame3, cliquesub, "")
                        dropsub.grid_forget()
                        dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                       activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                        dropsub.grid(row=1, column=3, sticky="w")

                    # \\\\\\\\\\\\\\\\\\\\\\\\\\\ FIM protheus \\\\\\\\\\\\\\\\\\\\\\\\\\\

                    def dropselecaotipo(event):
                        if clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Solicitação":
                            opt_protheus_solicitacao()
                        elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Problemas":
                            opt_protheus_problemas()
                        elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Dúvidas":
                            opt_protheus_duvidas()
                        elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Melhorias":
                            opt_protheus_melhorias()
                        elif clique_ocorr.get() == "Protheus" and cliquetipo.get() == "Projetos":
                            opt_protheus_projetos()
                        elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Solicitação":
                            opt_windows_solicitacao()
                        elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Problemas":
                            opt_windows_problemas()
                        elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Dúvidas":
                            opt_windows_duvidas()
                        elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Melhorias":
                            opt_windows_melhorias()
                        elif clique_ocorr.get() == "Windows" and cliquetipo.get() == "Projetos":
                            opt_windows_projetos()
                        elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Solicitação":
                            opt_softwares_solicitacao()
                        elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Problemas":
                            opt_softwares_problemas()
                        elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Dúvidas":
                            opt_softwares_duvidas()
                        elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Melhorias":
                            opt_softwares_melhorias()
                        elif clique_ocorr.get() == "Softwares" and cliquetipo.get() == "Projetos":
                            opt_softwares_projetos()
                        elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Solicitação":
                            opt_email_solicitacao()
                        elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Problemas":
                            opt_email_problemas()
                        elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Dúvidas":
                            opt_email_duvidas()
                        elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Melhorias":
                            opt_email_melhorias()
                        elif clique_ocorr.get() == "E-mail" and cliquetipo.get() == "Projetos":
                            opt_email_projetos()
                        elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Solicitação":
                            opt_impressora_solicitacao()
                        elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Problemas":
                            opt_impressora_problemas()
                        elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Dúvidas":
                            opt_impressora_duvidas()
                        elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Melhorias":
                            opt_impressora_melhorias()
                        elif clique_ocorr.get() == "Impressora" and cliquetipo.get() == "Projetos":
                            opt_impressora_projetos()
                        elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Solicitação":
                            opt_internet_solicitacao()
                        elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Problemas":
                            opt_internet_problemas()
                        elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Dúvidas":
                            opt_internet_duvidas()
                        elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Melhorias":
                            opt_internet_melhorias()
                        elif clique_ocorr.get() == "Internet" and cliquetipo.get() == "Projetos":
                            opt_internet_projetos()
                        elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Solicitação":
                            opt_radio_solicitacao()
                        elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Problemas":
                            opt_radio_problemas()
                        elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Dúvidas":
                            opt_radio_duvidas()
                        elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Melhorias":
                            opt_radio_melhorias()
                        elif clique_ocorr.get() == "Rádio" and cliquetipo.get() == "Projetos":
                            opt_radio_projetos()
                        elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Solicitação":
                            opt_rede_solicitacao()
                        elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Problemas":
                            opt_rede_problemas()
                        elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Dúvidas":
                            opt_rede_duvidas()
                        elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Melhorias":
                            opt_rede_melhorias()
                        elif clique_ocorr.get() == "Rede" and cliquetipo.get() == "Projetos":
                            opt_rede_projetos()
                        elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Solicitação":
                            opt_telefonia_solicitacao()
                        elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Problemas":
                            opt_telefonia_problemas()
                        elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Dúvidas":
                            opt_telefonia_duvidas()
                        elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Melhorias":
                            opt_telefonia_melhorias()
                        elif clique_ocorr.get() == "Telefonia" and cliquetipo.get() == "Projetos":
                            opt_telefonia_projetos()
                        elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Solicitação":
                            opt_hardware_solicitacao()
                        elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Problemas":
                            opt_hardware_problemas()
                        elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Dúvidas":
                            opt_hardware_duvidas()
                        elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Melhorias":
                            opt_hardware_melhorias()
                        elif clique_ocorr.get() == "Hardware" and cliquetipo.get() == "Projetos":
                            opt_hardware_projetos()
                        elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Solicitação":
                            opt_acessos_solicitacao()
                        elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Problemas":
                            opt_acessos_problemas()
                        elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Dúvidas":
                            opt_acessos_duvidas()
                        elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Melhorias":
                            opt_acessos_melhorias()
                        elif clique_ocorr.get() == "Acessos" and cliquetipo.get() == "Projetos":
                            opt_acessos_projetos()

                    def dropselecao_ocorr(event):
                        if clique_ocorr.get() == "Protheus":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")

                        elif clique_ocorr.get() == "Windows":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Softwares":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "E-mail":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Impressora":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Internet":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Rádio":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Rede":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Telefonia":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Hardware":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")
                        elif clique_ocorr.get() == "Acessos":
                            entrytitulo.config(state='normal')
                            entrytitulo.delete(0, END)
                            entrytitulo.config(state='disabled')
                            cliquetipo.set('')
                            droptipo.config(state=NORMAL, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                            activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE,
                                            width=24)
                            cliquesub = StringVar()
                            dropsub = OptionMenu(frame3, cliquesub, "")
                            dropsub.grid_forget()
                            dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                           activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                            dropsub.grid(row=1, column=3, sticky="w")

                    def anexo():
                        anexo = filedialog.askopenfilename(initialdir="os.path.expanduser(default_dir)",
                                                           title="Escolha um Arquivo",
                                                           filetypes=([("Todos os arquivos", "*.*")]), parent=root3)
                        entryanexo.config(state=NORMAL)
                        entryanexo.insert(0, anexo)
                        entryanexo.config(state=DISABLED)
                        l_nome_anexo = os.path.basename(anexo)
                        l_caminho_anexo = anexo
                        global nome_anexo
                        nome_anexo = l_nome_anexo
                        global caminho_anexo
                        caminho_anexo = l_caminho_anexo

                    def confirmar_edicao():
                        if cliquetipo.get() == "" or entrytitulo.get() == "" or txtdescr.get("1.0",'end-1c') == "":
                            messagebox.showwarning('+ Abrir Chamado: Erro',
                                                   'Todos os campos com ( * ) devem ser preenchidos.',
                                                   parent=root3)
                        else:
                            ocorr = clique_ocorr.get().upper()
                            tipo = cliquetipo.get().upper()
                            titulo = entrytitulo.get().upper()
                            anexo_nome = nome_anexo
                            descricao_problema = txtdescr.get("1.0", 'end-1c')
                            nome_maquina = entrynomemaquina.get().upper()
                            ramal = entryramal.get().upper()
                            setor = clique_setor.get().upper()
                            resposta = messagebox.askyesno('Atenção:',
                                                           f'Tem certeza de que deseja editar o chamado nº{result[0]}?',
                                                           parent=root3)

                            if resposta == True:
                                if entryanexo.get() == "":
                                    cursor.execute(
                                        "UPDATE helpdesk.dbo.chamados SET ocorrencia = ?, tipo = ?,  titulo = ?, descricao_problema = ?, nome_maquina = ?, ramal = ?, setor = ? WHERE id_chamado = ?",
                                        (ocorr, tipo, titulo, descricao_problema, nome_maquina, ramal, setor, result[0]))
                                    cursor.commit()
                                    atualizar_lista_principal()
                                    root3.destroy()
                                    root2.destroy()
                                else:
                                    subprocess.call(r'net use /delete \\192.168.1.19', shell=True)
                                    subprocess.call(r'net use \\192.168.1.19 /user:impressoras gv2K17ADM',
                                                    shell=True)
                                    fsize = os.stat(caminho_anexo)
                                    if fsize.st_size > 5242880:
                                        messagebox.showwarning('Erro:', 'Anexo superior a 5MB.', parent=root3)
                                    else:
                                        pasta = r'\\192.168.1.19/helpdesk/anexos/' + result[23]
                                        try:
                                            shutil.copy(caminho_anexo, pasta)
                                        except:
                                            messagebox.showwarning('Erro:',
                                                                   'Erro ao enviar o anexo para o diretório.',
                                                                   parent=root3)
                                            return False

                                        try:
                                            cursor.execute(
                                                "UPDATE helpdesk.dbo.chamados SET ocorrencia = ?, tipo = ?,  titulo = ?, descricao_problema = ?, nome_maquina = ?, ramal = ?, setor = ?, nome_anexo = ? WHERE id_chamado = ?",
                                                (ocorr, tipo, titulo, descricao_problema, nome_maquina, ramal, setor,
                                                 anexo_nome, result[0]))
                                            cursor.commit()
                                        except:
                                            messagebox.showwarning('Erro:',
                                                                   'Erro ao salvar as informações no banco de dados.',
                                                                   parent=root3)
                                            return False
                                    messagebox.showinfo('Sucesso:',
                                                           'Edição realizada com sucesso.',
                                                           parent=root3)
                                    atualizar_lista_principal()
                                    root3.destroy()
                                    root2.destroy()

                    def cancelar():
                        root2.destroy()

                    def setup_entradas():
                        entryusuario.insert(0, result[1])
                        entryusuario.config(state='disabled')
                        entry_solicitante.insert(0, result[16])
                        entry_solicitante.config(state='disabled')
                        clique_ocorr.set(result[17])
                        cliquetipo.set(result[4])
                        clique_setor.set(result[10])

                        entrytitulo.config(state='normal')
                        entrytitulo.insert(0, result[5])
                        txtdescr.insert(END, result[7])
                        entrynomemaquina.insert(0, result[8])
                        entryramal.insert(0, result[9])
                        entrysetor.insert(0, result[10])

                        if result[21] != None:
                            entryemail.config(state='normal')
                            entryemail.insert(0, result[21])
                            entryemail.config(state='disabled')

                    frame0 = Frame(root3, bg='#ffffff')
                    frame0.grid(row=0, column=0, stick='nsew')
                    root2.grid_rowconfigure(0, weight=1)
                    root2.grid_columnconfigure(0, weight=1)
                    frame1 = Frame(frame0, bg="#1d366c")
                    frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
                    frame2 = Frame(frame0, bg='#ffffff')
                    frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                    frame3 = Frame(frame0, bg='#ffffff')
                    frame3.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                    frame4 = Frame(frame0, bg='#ffffff')
                    frame4.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                    frame5 = Frame(frame0, bg='#ffffff')
                    frame5.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                    frame6 = Frame(frame0, bg='#ffffff')
                    frame6.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                    frame7 = Frame(frame0, bg='#ffffff')
                    frame7.pack(side=TOP, fill=X, expand=False, anchor='center', pady=8)
                    frame8 = Frame(frame0, bg='#1d366c')
                    frame8.pack(side=TOP, fill=X, expand=False, anchor='center')

                    Label(frame1, image=nova_image_chamado, text=f" Editar Chamado: Nº {result[0]}",
                          compound="left",
                          bg='#1d366c',
                          fg='#FFFFFF',
                          font=fonte_titulos).grid(row=0, column=1)
                    frame1.grid_columnconfigure(0, weight=1)
                    frame1.grid_columnconfigure(2, weight=1)

                    Label(frame2, text="Usuário Logado:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1,
                                                                                                sticky="w")
                    entryusuario = Entry(frame2, font=fonte_padrao, justify='center')
                    entryusuario.grid(row=1, column=1, sticky="w")

                    Label(frame2, text="Solicitante: *", font=fonte_padrao, fg='#8B0000', bg='#ffffff').grid(row=0,
                                                                                                             column=2,
                                                                                                             sticky="w",
                                                                                                             padx=12)
                    entry_solicitante = Entry(frame2, font=fonte_padrao, justify='center')
                    entry_solicitante.grid(row=1, column=2, sticky="w", padx=12)
                    entry_solicitante.focus_force()

                    Label(frame2, text="Data de Abertura:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=3,
                                                                                                  sticky="w",
                                                                                                  padx=12)
                    entrydataabertura = Entry(frame2, font=fonte_padrao, justify='center')
                    entrydataabertura.grid(row=1, column=3, sticky="w", padx=12)
                    entrydataabertura.insert(0, data)
                    entrydataabertura.config(state='disabled')

                    Label(frame2, text="E-mail:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=4,
                                                                                                  sticky="w")
                    entryemail = Entry(frame2, font=fonte_padrao, justify='center', width=36)
                    entryemail.grid(row=1, column=4, sticky="w")
                    entryemail.config(state='disabled')
                    frame2.grid_columnconfigure(0, weight=1)
                    frame2.grid_columnconfigure(5, weight=1)

                    Label(frame3, text="Ocorrência: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0,
                                                                                                            column=1,
                                                                                                            sticky="w")
                    clique_ocorr = StringVar()
                    drop_ocorr = OptionMenu(frame3, clique_ocorr, "Protheus", "Windows", "Softwares", "E-mail",
                                            "Impressora", "Internet", "Rádio", "Rede", "Telefonia", "Hardware", "Acessos",
                                            command=dropselecao_ocorr)
                    drop_ocorr.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                      activeforeground="#FFFFFF",
                                      highlightthickness=0, relief=RIDGE, width=24, cursor="hand2")
                    drop_ocorr.grid(row=1, column=1, sticky="w")

                    Label(frame3, text="Tipo: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0,
                                                                                                      column=2,
                                                                                                      sticky="w",
                                                                                                      padx=10)
                    cliquetipo = StringVar()
                    droptipo = OptionMenu(frame3, cliquetipo, "Solicitação", "Problemas", "Dúvidas", "Melhorias",
                                          "Projetos", command=dropselecaotipo)
                    droptipo.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                    activeforeground="#FFFFFF",
                                    highlightthickness=0, relief=RIDGE, width=24, cursor="hand2")
                    droptipo.grid(row=1, column=2, sticky="w", padx=10)
                    Label(frame3, text="Título Predefinido: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(
                        row=0,
                        column=3,
                        sticky="w")
                    cliquesub = StringVar()
                    global dropsub
                    dropsub = OptionMenu(frame3, cliquesub, "")
                    dropsub.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                                   activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=51, cursor="hand2")
                    dropsub.grid(row=1, column=3, sticky="w")
                    frame3.grid_columnconfigure(0, weight=1)
                    frame3.grid_columnconfigure(4, weight=1)

                    image_anexo = Image.open('imagens\\anexo.png')
                    resize_anexo = image_anexo.resize((15, 20))
                    nova_image_anexo = ImageTk.PhotoImage(resize_anexo)
                    btnanexo = Button(frame4, image=nova_image_anexo, text=" Anexar arquivo.", compound="left",
                                      font=fonte_padrao, bg='#ffffff', fg='#1d366c', command=anexo,
                                      borderwidth=0, relief=RIDGE, activebackground="#ffffff",
                                      activeforeground="#7c7c7c", cursor="hand2")
                    btnanexo.photo = nova_image_anexo
                    btnanexo.grid(row=0, column=2, sticky="w", padx=(9, 0))

                    lbltitulo = Label(frame4, text="Título: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000')
                    lbltitulo.grid(row=0, column=1, sticky="w")
                    entrytitulo = Entry(frame4, font=fonte_padrao, justify='center', width=66)
                    entrytitulo.grid(row=1, column=1, sticky="ew", padx=(0, 9))
                    entrytitulo.config(state=DISABLED)
                    entryanexo = Entry(frame4, font=fonte_padrao, justify='center', width=36)
                    entryanexo.grid(row=1, column=2, sticky="ew", padx=(9, 0))
                    entryanexo.config(state='disabled')
                    frame4.grid_columnconfigure(0, weight=1)
                    frame4.grid_columnconfigure(3, weight=1)

                    Label(frame5, text="Descrição do Problema: *", font=fonte_padrao, bg='#ffffff',
                          fg='#8B0000').grid(
                        row=0, column=1, sticky="w")
                    txtdescr = scrolledtext.ScrolledText(frame5, width=103, height=10, font=fonte_padrao)
                    txtdescr.grid(row=1, column=1)
                    frame5.grid_columnconfigure(0, weight=1)
                    frame5.grid_columnconfigure(2, weight=1)

                    Label(frame6, text="Nome da Máquina:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(
                        row=0,
                        column=1,
                        sticky="w")
                    entrynomemaquina = Entry(frame6, font=fonte_padrao, justify='center', width=25)
                    entrynomemaquina.grid(row=1, column=1, sticky="w")

                    Label(frame6, text="Ramal:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w",
                                                                                       padx=14)
                    entryramal = Entry(frame6, font=fonte_padrao, justify='center', width=25)
                    entryramal.grid(row=1, column=2, sticky="w", padx=14)

                    Label(frame6, text="Setor: *", font=fonte_padrao, bg='#ffffff', fg='#8B0000').grid(row=0,
                                                                                                       column=3,sticky="w")
                    OptionList = [
                        "Aciaria",
                        "Almoxarifado",
                        "Ambulatório",
                        "Balança",
                        "Comercial",
                        "Compras",
                        "Contabilidade",
                        "Custos",
                        "EHS",
                        "Elétrica",
                        "Engenharia",
                        "Faturamento",
                        "Financeiro",
                        "Fiscal",
                        "Lab Inspeção",
                        "Lab Mecânico",
                        "Lab Químico",
                        "Laminação",
                        "Logística",
                        "Oficina de Cilindros",
                        "Oficina Mecânica",
                        "Pátio de Sucata",
                        "PCP",
                        "Planta D´agua",
                        "Planta de Escória",
                        "Portaria",
                        "Qualidade",
                        "Refratários",
                        "RH",
                        "Subestação",
                        "TI",
                        "Utilidades"
                    ]
                    clique_setor = StringVar()
                    drop_setor = OptionMenu(frame6, clique_setor, *OptionList)
                    drop_setor.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                      activeforeground="#FFFFFF",
                                      highlightthickness=0, relief=RIDGE, width=25, cursor="hand2")
                    drop_setor.grid(row=1, column=3, sticky="w")

                    frame6.grid_columnconfigure(0, weight=1)
                    frame6.grid_columnconfigure(4, weight=1)

                    bt1 = Button(frame7, text='Confirmar', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                                 activeforeground="#FFFFFF", highlightthickness=0, width=10, relief=RIDGE,
                                 command=confirmar_edicao,
                                 font=fonte_padrao, cursor="hand2")
                    bt1.grid(row=0, column=1, padx=5)
                    bt2 = Button(frame7, text='Cancelar', width=10, relief=RIDGE, command=cancelar,
                                 font=fonte_padrao, cursor="hand2")
                    bt2.grid(row=0, column=2, padx=5)
                    frame7.grid_columnconfigure(0, weight=1)
                    frame7.grid_columnconfigure(3, weight=1)

                    Label(frame8, text=" ", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos).grid(row=0, column=1)
                    frame8.grid_columnconfigure(0, weight=1)
                    frame8.grid_columnconfigure(2, weight=1)
                    setup_entradas()
                    '''root2.update()
                    largura = frame0.winfo_width()
                    altura = frame0.winfo_height()
                    print(largura, altura)'''
                    window_width = 742
                    window_height = 577
                    screen_width = root2.winfo_screenwidth()
                    screen_height = root2.winfo_screenheight()
                    x_cordinate = int((screen_width / 2) - (window_width / 2))
                    y_cordinate = int((screen_height / 2) - (window_height / 2))
                    root3.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
                    root3.resizable(0, 0)
                    root3.configure(bg='#000000')
                    root3.iconbitmap('imagens\\ico.ico')

            frame0 = Frame(root2, bg='#ffffff')
            frame0.grid(row=0, column=1, sticky=NSEW)
            frame_topo = Frame(frame0, bg='#1d366c')
            frame_topo.grid(row=0, column=1, sticky=NSEW, columnspan=2)
            frame_esquerda = LabelFrame(frame0, text=f"Informações do Chamado -- Aberto por: {result[1]}",
                                        font=fonte_padrao, bg='#ffffff')
            frame_esquerda.grid(row=1, column=1, sticky=NSEW, padx=10)
            frame_direita = Frame(frame0, bg='#ffffff')
            frame_direita.grid(row=1, column=2, sticky=NSEW, padx=10)
            frame_baixo = Frame(frame0, bg='#1d366c')
            frame_baixo.grid(row=2, column=1, sticky=NSEW, columnspan=2)
            # //// TOPO ////
            Label(frame_topo, image=nova_image_visualizarchamado, text=f" Visualizando Chamado: Nº {n_chamado}",
                  compound="left", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos).grid(row=0, column=1)
            frame_topo.grid_columnconfigure(0, weight=1)
            frame_topo.grid_columnconfigure(2, weight=1)
            # //// ESQUERDA ////
            frame1 = Frame(frame_esquerda, bg='#ffffff')
            frame1.grid(row=0, column=1, sticky=EW, pady=(6, 0))
            frame2 = Frame(frame_esquerda, bg='#ffffff')
            frame2.grid(row=1, column=1, sticky=NSEW, pady=6)
            frame3 = Frame(frame_esquerda, bg='#ffffff')
            frame3.grid(row=2, column=1, sticky=NSEW, pady=6)
            frame4 = Frame(frame_esquerda, bg='#ffffff')
            frame4.grid(row=3, column=1, sticky=NSEW, pady=6)
            frame5 = Frame(frame_esquerda, bg='#ffffff')
            frame5.grid(row=4, column=1, sticky=NSEW, pady=6)
            frame6 = Frame(frame_esquerda, bg='#ffffff')
            frame6.grid(row=5, column=1, sticky=NSEW, pady=6)

            Label(frame1, text="Solicitante:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w",
                                                                                     padx=6)
            entrysolicitante = Entry(frame1, font=fonte_padrao, justify='center', bg='#ffffff')
            entrysolicitante.grid(row=1, column=1, sticky="w", padx=6)

            Label(frame1, text="Data de Abertura:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w",
                                                                                          padx=6)
            entrydataabertura = Entry(frame1, font=fonte_padrao, justify='center', bg='#ffffff')
            entrydataabertura.grid(row=1, column=2, sticky="w", padx=6)

            Label(frame1, text="E-mail:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=3, sticky="w", padx=6)
            entryemail = Entry(frame1, font=fonte_padrao, justify='center', bg='#ffffff', width=34)
            entryemail.grid(row=1, column=3, sticky="w", padx=6)

            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(4, weight=1)

            Label(frame2, text="Hora:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w", padx=6)
            entryhoraabertura = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=10)
            entryhoraabertura.grid(row=1, column=1, sticky="w", padx=6)

            Label(frame2, text="Nome da Máquina:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w",
                                                                                          padx=6)
            entrynomemaquina = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=20)
            entrynomemaquina.grid(row=1, column=2, sticky="w", padx=6)

            Label(frame2, text="Ramal:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=3, sticky="w", padx=6)
            entryramal = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=15)
            entryramal.grid(row=1, column=3, sticky="w", padx=6)

            Label(frame2, text="Setor:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=4, sticky="w", padx=6)
            entrysetor = Entry(frame2, font=fonte_padrao, justify='center', bg='#ffffff', width=27)
            entrysetor.grid(row=1, column=4, sticky="w", padx=6)

            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(5, weight=1)

            Label(frame3, text="Ocorrência:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w", padx=6)
            entryocorrencia = Entry(frame3, font=fonte_padrao, justify='center', bg='#ffffff', width=26)
            entryocorrencia.grid(row=1, column=1, sticky="ew", padx=6)

            Label(frame3, text="Tipo:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=2, sticky="w", padx=6)
            entrytipo = Entry(frame3, font=fonte_padrao, justify='center', bg='#ffffff', width=50)
            entrytipo.grid(row=1, column=2, columnspan=2, sticky="w", padx=6)

            frame3.grid_columnconfigure(0, weight=1)
            frame3.grid_columnconfigure(3, weight=1)

            lbltitulo = Label(frame4, text="Título:*", font=fonte_padrao, bg='#ffffff')
            lbltitulo.grid(row=0, column=1, sticky="w")
            entrytitulo = Entry(frame4, font=fonte_padrao, justify='center', width=79)
            entrytitulo.grid(row=1, column=1, sticky="w", columnspan=3)

            frame4.grid_columnconfigure(0, weight=1)
            frame4.grid_columnconfigure(3, weight=1)

            Label(frame5, text="Descrição do Problema:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1,
                                                                                                sticky="w")
            txtdescr_atendimento = scrolledtext.ScrolledText(frame5, width=77, height=16, font=fonte_padrao,
                                                             bg='#ffffff', wrap=WORD)
            txtdescr_atendimento.grid(row=1, column=1, sticky="ew")

            frame5.grid_columnconfigure(0, weight=1)
            frame5.grid_columnconfigure(2, weight=1)

            def muda_anexo(e):
                image_abriranexo = Image.open('imagens\\anexo_over.png')
                resize_abriranexo = image_abriranexo.resize((15, 20))
                nova_image_abriranexo = ImageTk.PhotoImage(resize_abriranexo)
                btabrir_anexo.photo = nova_image_abriranexo
                btabrir_anexo.config(image=nova_image_abriranexo, fg='#7c7c7c')

            def volta_anexo(e):
                image_abriranexo = Image.open('imagens\\anexo.png')
                resize_abriranexo = image_abriranexo.resize((15, 20))
                nova_image_abriranexo = ImageTk.PhotoImage(resize_abriranexo)
                btabrir_anexo.photo = nova_image_abriranexo
                btabrir_anexo.config(image=nova_image_abriranexo, fg='#1d366c')

            image_abriranexo = Image.open('imagens\\anexo.png')
            resize_abriranexo = image_abriranexo.resize((15, 20))
            nova_image_abriranexo = ImageTk.PhotoImage(resize_abriranexo)
            btabrir_anexo = Button(frame6, image=nova_image_abriranexo, text=" Abrir Anexo.", compound="left",
                                   font=fonte_padrao, bg='#ffffff', fg='#1d366c', command=abrir_anexo,
                                   borderwidth=0, relief=RIDGE, activebackground="#ffffff",
                                   activeforeground="#7c7c7c", cursor="hand2")
            btabrir_anexo.photo = nova_image_abriranexo
            btabrir_anexo.grid(row=0, column=0, padx=(6, 0))
            btabrir_anexo.bind("<Enter>", muda_anexo)
            btabrir_anexo.bind("<Leave>", volta_anexo)

            bt_editar_chamado = Button(frame6, text='Editar', width=10, relief=RIDGE, command=editar_chamado,
                                       font=fonte_padrao, bg="#1d366c", fg="#ffffff",
                                       activebackground="#1d366c",
                                       activeforeground="#FFFFFF", state=NORMAL, cursor="hand2")
            bt_editar_chamado.grid(row=0, column=2, padx=(0, 6))
            bt_excluir_chamado = Button(frame6, text='Excluir', command=excluir_chamado, width=10, font=fonte_padrao, cursor="hand2")
            bt_excluir_chamado.grid(row=0, column=3, padx=6)
            frame6.grid_columnconfigure(1, weight=1)
            frame6.grid_columnconfigure(1, weight=1)

            # //// DIREITA ////
            frame1 = Frame(frame_direita, bg='#ffffff')
            frame1.grid(row=0, column=1, sticky=EW)
            frame2 = Frame(frame_direita, bg='#ffffff')
            frame2.grid(row=1, column=1, sticky=NSEW, pady=6)
            frame3 = LabelFrame(frame_direita, text="Interação do Chamado (Usuário\Analistas):", font=fonte_padrao,
                                bg='#ffffff')
            frame3.grid(row=2, column=1, sticky=NSEW, pady=6)
            frame4 = Frame(frame_direita, bg='#ffffff')
            frame4.grid(row=3, column=1, sticky=NSEW)
            frame5 = Frame(frame_direita, bg='#ffffff')
            frame5.grid(row=4, column=1, sticky=NSEW)

            Label(frame1, text="Nome do Analista:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w")
            entryanalista2 = Entry(frame1, font=fonte_padrao, justify='center', width=70, bg='#ffffff')
            entryanalista2.grid(row=1, column=1, sticky="ew")
            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(2, weight=1)

            Label(frame2, text="Prioridade do Chamado:", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1,
                                                                                               sticky="w", columnspan=5)
            options = ["Baixa", "Média", "Alta", "Urgente"]
            clique_prioridade = StringVar()

            drop_prioridade = OptionMenu(frame2, clique_prioridade, *options)
            drop_prioridade.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                               highlightthickness=0, relief=RIDGE, width=20, cursor="hand2")
            drop_prioridade.grid(row=1, column=1, columnspan=5, sticky="w", pady=(0, 10))

            Label(frame2, text="Status:", font=fonte_padrao, bg='#ffffff').grid(row=2, column=1, sticky="w")
            options = ["Em andamento", "Encerrado"]
            clique_status = StringVar()

            drop_status = OptionMenu(frame2, clique_status, *options)
            drop_status.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                               highlightthickness=0, relief=RIDGE, width=20, cursor="hand2")
            drop_status.grid(row=3, column=1, sticky="w")
            Label(frame2, text="Data do Atendimento:", font=fonte_padrao, bg='#ffffff').grid(row=2, column=2)

            def muda_calendario_atendimento(e):
                image_calendario_atendimento = Image.open('imagens\\agenda_over.png')
                resize_calendario_atendimento = image_calendario_atendimento.resize((22, 22))
                nova_image_calendario_atendimento = ImageTk.PhotoImage(resize_calendario_atendimento)
                btn_calendario_atendimento.photo = nova_image_calendario_atendimento
                btn_calendario_atendimento.config(image=nova_image_calendario_atendimento, fg='#7c7c7c')

            def volta_calendario_atendimento(e):
                image_calendario_atendimento = Image.open('imagens\\agenda.png')
                resize_calendario_atendimento = image_calendario_atendimento.resize((22, 22))
                nova_image_calendario_atendimento = ImageTk.PhotoImage(resize_calendario_atendimento)
                btn_calendario_atendimento.photo = nova_image_calendario_atendimento
                btn_calendario_atendimento.config(image=nova_image_calendario_atendimento, fg='#1d366c')

            image_calendario_atendimento = Image.open('imagens\\agenda.png')
            resize_calendario_atendimento = image_calendario_atendimento.resize((22, 22))
            nova_image_calendario_atendimento = ImageTk.PhotoImage(resize_calendario_atendimento)
            btn_calendario_atendimento = Button(frame2, image=nova_image_calendario_atendimento, font=fonte_padrao,
                                                bg='#ffffff', fg='#1d366c', borderwidth=0, relief=RIDGE, activebackground="#ffffff",
                                                activeforeground="#7c7c7c", cursor="hand2")
            btn_calendario_atendimento.photo = nova_image_calendario_atendimento
            btn_calendario_atendimento.grid(row=2, column=3, sticky="w")
            btn_calendario_atendimento.bind("<Enter>", muda_calendario_atendimento)
            btn_calendario_atendimento.bind("<Leave>", volta_calendario_atendimento)
            entrydataatendimento2 = Entry(frame2, font=fonte_padrao, justify='center', width=22, bg='#ffffff')
            entrydataatendimento2.grid(row=3, column=2, columnspan=2, sticky="we", padx=8)
            entrydataatendimento2.config(state='disabled')
            Label(frame2, text="Data de Encerramento:", font=fonte_padrao, bg='#ffffff').grid(row=2, column=4,
                                                                                              sticky="w")

            def muda_calendario_encerramento(e):
                image_calendario_encerramento = Image.open('imagens\\agenda_over.png')
                resize_calendario_encerramento = image_calendario_encerramento.resize((22, 22))
                nova_image_calendario_encerramento = ImageTk.PhotoImage(resize_calendario_encerramento)
                btn_calendario_encerramento.photo = nova_image_calendario_encerramento
                btn_calendario_encerramento.config(image=nova_image_calendario_encerramento, fg='#7c7c7c')

            def volta_calendario_encerramento(e):
                image_calendario_encerramento = Image.open('imagens\\agenda.png')
                resize_calendario_encerramento = image_calendario_encerramento.resize((22, 22))
                nova_image_calendario_encerramento = ImageTk.PhotoImage(resize_calendario_encerramento)
                btn_calendario_encerramento.photo = nova_image_calendario_encerramento
                btn_calendario_encerramento.config(image=nova_image_calendario_encerramento, fg='#1d366c')

            image_calendario_encerramento = Image.open('imagens\\agenda.png')
            resize_calendario_encerramento = image_calendario_encerramento.resize((22, 22))
            nova_image_calendario_encerramento = ImageTk.PhotoImage(resize_calendario_encerramento)
            btn_calendario_encerramento = Button(frame2, image=nova_image_calendario_encerramento, borderwidth=0, relief=RIDGE,
                                                 activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btn_calendario_encerramento.photo = nova_image_calendario_encerramento
            btn_calendario_encerramento.grid(row=2, column=5, sticky="w")
            btn_calendario_encerramento.bind("<Enter>", muda_calendario_encerramento)
            btn_calendario_encerramento.bind("<Leave>", volta_calendario_encerramento)
            entrydataencerramento = Entry(frame2, font=fonte_padrao, justify='center', width=22, bg='#ffffff')
            entrydataencerramento.grid(row=3, column=4, columnspan=2, sticky="we")
            entrydataencerramento.config(state='disabled')
            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(5, weight=1)

            entryinteracao = Entry(frame3, font=fonte_padrao, justify='center', width=46, bg='#ffffff')
            entryinteracao.grid(row=0, column=1, padx=(2, 0))
            entryinteracao.focus_force()

            btn_envia_interacao = Button(frame3, text='Enviar', width=10, relief=RIDGE, font=fonte_padrao,
                                            command=enviar_int, bg="#1d366c", fg="#ffffff", activebackground="#1d366c",
                                            activeforeground="#FFFFFF", state=NORMAL, cursor="hand2")
            btn_envia_interacao.grid(row=0, column=2, padx=(5, 0))
            btn_cancela_interacao = Button(frame3, text='Desfazer', width=10, font=fonte_padrao, command=desfazer_int, cursor="hand2")
            btn_cancela_interacao.grid(row=0, column=3, padx=(5, 2))

            Label(frame3, text="Histórico:", font=fonte_padrao, bg='#ffffff').grid(row=1, column=1, sticky="w")
            txtinteracao = scrolledtext.ScrolledText(frame3, font=fonte_padrao, width=68, height=10, bg='#ffffff', wrap=WORD)
            txtinteracao.grid(row=2, column=1, columnspan=3, sticky="ew", padx=2, pady=(0, 2))
            txtinteracao.config(state='disabled', fg='#696969')

            frame3.grid_columnconfigure(0, weight=1)
            frame3.grid_columnconfigure(4, weight=1)

            Label(frame4, text="Solução:*", font=fonte_padrao, bg='#ffffff').grid(row=0, column=1, sticky="w")
            txtsolucao = scrolledtext.ScrolledText(frame4, width=71, height=5, font=fonte_padrao, bg='#ffffff', wrap=WORD)
            txtsolucao.grid(row=1, column=1, sticky="ew", columnspan=2)
            txtsolucao.config(state='disabled', fg='#696969')
            frame4.grid_columnconfigure(0, weight=1)
            frame4.grid_columnconfigure(3, weight=1)

            btcancelar_atendimento = Button(frame5, text='Salvar', width=10, relief=RIDGE, command=salvar,
                                            font=fonte_padrao, bg="#1d366c", fg="#ffffff", activebackground="#1d366c",
                                            activeforeground="#FFFFFF", state=NORMAL, cursor="hand2")
            btcancelar_atendimento.grid(row=0, column=1, padx=5, pady=6)
            btconfirma_atencimento = Button(frame5, text='Cancelar', command=cancelar, width=10, font=fonte_padrao, cursor="hand2")
            btconfirma_atencimento.grid(row=0, column=2, padx=5, pady=6)

            frame5.grid_columnconfigure(0, weight=1)
            frame5.grid_columnconfigure(3, weight=1)

            # //// BAIXO ////
            Label(frame_baixo, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos).grid(row=0, column=1)
            frame_baixo.grid_columnconfigure(0, weight=1)
            frame_baixo.grid_columnconfigure(2, weight=1)
            setup_entradas()
            '''root2.update()
            largura = frame0.winfo_width()
            altura = frame0.winfo_height()
            print(largura, altura)'''
            window_width = 1136
            window_height = 647
            screen_width = root2.winfo_screenwidth()
            screen_height = root2.winfo_screenheight()
            x_cordinate = int((screen_width / 2) - (window_width / 2))
            y_cordinate = int((screen_height / 2) - (window_height / 2))
            root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
            root2.resizable(0, 0)
            root2.configure(bg='#ffffff')
            root2.iconbitmap('imagens\\ico.ico')
            root2.grid_columnconfigure(0, weight=1)
            root2.grid_columnconfigure(2, weight=1)
    # /////////////////////////////FIM VISUALIZAR/////////////////////////////

    def pesquisar_bind(event):
        pesquisar()

    def pesquisar():
        # Parar a atualização automática da lista principal
        global controle_loop
        controle_loop = 1
        global ativa_filtro
        ativa_filtro = 1
        filtro = ent_busca.get()
        global pesquisa_com_filtro_tabela
        pesquisa_com_filtro_tabela = clique_busca.get()
        global pesquisa_com_filtro_filtro
        pesquisa_com_filtro_filtro = filtro

        if clique_busca.get() == "Título":
            if nivel_acesso == 0:
                cursor.execute(
                    "SELECT * FROM dbo.chamados WHERE titulo LIKE '%' + ? + '%' AND solicitante =? ORDER BY id_chamado DESC",
                    (filtro, usuariologado,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute(
                    "SELECT * FROM dbo.chamados WHERE titulo LIKE '%' + ? + '%' AND solicitante =? ORDER BY id_chamado DESC",
                    (filtro, usuariologado,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))
            else:
                cursor.execute("SELECT * FROM dbo.chamados WHERE titulo LIKE '%' + ? + '%' ORDER BY id_chamado DESC",
                               (filtro,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.chamados WHERE titulo LIKE '%' + ? + '%' ORDER BY id_chamado DESC",
                                   (filtro,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))

        if clique_busca.get() == "Status":
            if nivel_acesso == 0:
                cursor.execute(
                    "SELECT * FROM dbo.chamados WHERE status LIKE '%' + ? + '%' AND solicitante =? ORDER BY id_chamado",
                    (filtro, usuariologado,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))
            else:
                cursor.execute("SELECT * FROM dbo.chamados WHERE status LIKE '%' + ? + '%' ORDER BY id_chamado",
                               (filtro,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.chamados WHERE status LIKE '%' + ? + '%' ORDER BY id_chamado",
                                   (filtro,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))

        elif clique_busca.get() == "Nº Chamado":
            if nivel_acesso == 0:
                cursor.execute("SELECT * FROM dbo.chamados WHERE id_chamado =? AND solicitante =?",
                               (filtro, usuariologado,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.chamados WHERE id_chamado =? AND solicitante =?",
                                   (filtro, usuariologado,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))
            else:
                cursor.execute("SELECT * FROM dbo.chamados WHERE id_chamado =?", (filtro,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.chamados WHERE id_chamado =?", (filtro,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))

        elif clique_busca.get() == "Solicitante":
            if nivel_acesso == 0:
                messagebox.showwarning('Erro:',
                                       'Não é possível efetuar uma busca pelo próprio nome ou de outro usuário.',
                                       parent=root)
            else:
                cursor.execute("SELECT * FROM dbo.chamados WHERE solicitante LIKE '%' + ? + '%' ORDER BY status",
                               (filtro,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.chamados WHERE solicitante LIKE '%' + ? + '%' ORDER BY status",
                                   (filtro,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))

        elif clique_busca.get() == "Ocorrência":
            if nivel_acesso == 0:
                cursor.execute(
                    "SELECT * FROM dbo.chamados WHERE ocorrencia LIKE '%' + ? + '%' AND solicitante =? ORDER BY status, id_chamado DESC",
                    (filtro, usuariologado,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute(
                        "SELECT * FROM dbo.chamados WHERE ocorrencia LIKE '%' + ? + '%' AND solicitante =? ORDER BY status, id_chamado DESC",
                        (filtro, usuariologado,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))
            else:
                cursor.execute(
                    "SELECT * FROM dbo.chamados WHERE ocorrencia LIKE '%' + ? + '%' ORDER BY status, id_chamado",
                    (filtro,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute(
                        "SELECT * FROM dbo.chamados WHERE ocorrencia LIKE '%' + ? + '%' ORDER BY status, id_chamado",
                        (filtro,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))

        elif clique_busca.get() == "Analista":
            if nivel_acesso == 0:
                cursor.execute(
                    "SELECT * FROM dbo.chamados WHERE id_analista LIKE '%' + ? + '%' AND solicitante =? ORDER BY status, id_chamado DESC",
                    (filtro, usuariologado,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute(
                        "SELECT * FROM dbo.chamados WHERE id_analista LIKE '%' + ? + '%' AND solicitante =? ORDER BY status, id_chamado DESC",
                        (filtro, usuariologado,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))
            else:
                cursor.execute(
                    "SELECT * FROM dbo.chamados WHERE id_analista LIKE '%' + ? + '%' ORDER BY status, id_chamado DESC",
                    (filtro,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute(
                        "SELECT * FROM dbo.chamados WHERE id_analista LIKE '%' + ? + '%' ORDER BY status, id_chamado",
                        (filtro,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                              row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11],
                                              row[15]),
                                              tags=('par',))

        elif clique_busca.get() == "Data Encerramento":
            if nivel_acesso == 0:
                cursor.execute("SELECT * FROM dbo.chamados WHERE data_encerramento LIKE '%' + ? + '%' AND solicitante =? ORDER BY id_chamado DESC",
                               (filtro, usuariologado,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.chamados WHERE data_encerramento LIKE '%' + ? + '%' AND solicitante =? ORDER BY id_chamado DESC",(filtro, usuariologado,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                              tags=('par',))
            else:
                cursor.execute("SELECT * FROM dbo.chamados WHERE data_encerramento LIKE '%' + ? + '%' ORDER BY id_chamado DESC",(filtro,))
                busca = cursor.fetchone()
                if busca is None:
                    messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root)
                else:
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute(
                        "SELECT * FROM dbo.chamados WHERE data_encerramento LIKE '%' + ? + '%' ORDER BY id_chamado DESC",
                        (filtro,))
                    for row in cursor:
                        if row[13] == None:
                            row[13] = ''
                        if row[15] == None:
                            row[15] = ''
                        tree_principal.insert('', 'end', text=" ",
                                              values=(row[0], row[2], row[16], row[4], row[17], row[5], row[13], row[11], row[15]),
                                              tags=('par',))

    def drop_selecao_busca(event):
        # ///////////////////REMOVE BUSCA///////////////////
        if clique_busca.get() == "Remover Filtro":
            global controle_loop
            controle_loop = 0
            global ativa_filtro
            ativa_filtro = 0
            atualizar_lista_principal()
    def ferramentas_bind(event):
        if nivel_acesso == 0:
            messagebox.showwarning('Atenção:', 'Módulo (Ferramentas) bloqueado.', parent=root)
        else:
            ferramentas()
    def ferramentas():
        def permissao():
            if usuariologado == 'Administrador':
                btntrocasenha.config(state='disabled')
                btnestoque.config(state='disabled')
        def trocasenha():
            root3 = Toplevel()
            root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
            root3.unbind_class("Button", "<Key-space>")
            root3.focus_force()
            root3.grab_set()

            def confirmar_bind(event):
                confirmar()
            def confirmar():
                senha_atual = esenha_antiga.get()
                senha_nova = esenha_nova.get()
                senha_confirma = esenha_confirma.get()
                if senha_atual == "" or senha_nova == "" or senha_confirma == "":
                    messagebox.showwarning('Alterar Senha:','Preencha todos os campos!', parent=root3)
                else:
                    r = cursor.execute("SELECT * FROM helpdesk.dbo.analista WHERE nome_analista=?", (usuariologado,))
                    for login in r.fetchall():
                        filtro_pwd = login[3]

                    if bcrypt.checkpw(senha_atual.encode("utf-8"), filtro_pwd.encode("utf-8")):
                        print(filtro_pwd)

                        if senha_nova != senha_confirma :
                            messagebox.showwarning('Alterar Senha:', 'As senhas não conferem!', parent=root3)
                        elif senha_atual == senha_confirma:
                            messagebox.showwarning('Alterar Senha:', 'Escolhe uma senha diferente da atual.', parent=root3)
                        else:
                            hashed = bcrypt.hashpw(senha_confirma.encode("utf-8"), bcrypt.gensalt())
                            try:
                                cursor.execute(
                                "UPDATE helpdesk.dbo.analista SET senha = ? WHERE id_useranalista = ?",
                                (hashed.decode("utf-8"), login[0]))
                                cursor.commit()
                            except:
                                messagebox.showerror('Alterar Senha:', 'Erro de conexão com o Banco de Dados.',
                                                     parent=root3)
                                return False
                            messagebox.showinfo('Alterar Senha:', 'Senha atualizada com sucesso.', parent=root3)
                            root3.destroy()
                    else:
                        messagebox.showwarning('Alterar Senha:', 'Senha atual incorreta!', parent=root3)
            def sair():
                root3.destroy()




            frame0 = Frame(root3, bg='#ffffff')
            frame0.grid(row=0, column=0, stick='nsew')
            root3.grid_rowconfigure(0, weight=1)
            root3.grid_columnconfigure(0, weight=1)
            frame1 = Frame(frame0, bg="#1d366c")
            frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
            frame2 = Frame(frame0, bg='#ffffff')
            frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=10)
            frame3 = Frame(frame0, bg='#ffffff')
            frame3.pack(side=TOP, fill=X, expand=False, anchor='center')
            frame4 = Frame(frame0, bg='#1d366c')
            frame4.pack(side=TOP, fill=X, expand=True, anchor='center')

            image_trocasenha2 = Image.open('imagens\\trocasenha2.png')
            resize_trocasenha2 = image_trocasenha2.resize((35, 35))
            nova_image_trocasenha2 = ImageTk.PhotoImage(resize_trocasenha2)

            lbllogin = Label(frame1, image=nova_image_trocasenha2, text=" Alterar Senha", compound="left",
                             bg='#1d366c',
                             fg='#FFFFFF', font=fonte_titulos)
            lbllogin.photo = nova_image_trocasenha2
            lbllogin.grid(row=0, column=1)
            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(2, weight=1)

            Label(frame2, text="Senha atual:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=0, column=1,
                                                                                             sticky="ew")

            esenha_antiga = Entry(frame2, show="*", width=20, font=fonte_padrao)
            esenha_antiga.grid(row=0, column=2, sticky="w", padx=5, pady=10)
            esenha_antiga.focus_force()
            esenha_antiga.bind("<Return>", confirmar_bind)

            Label(frame2, text="Nova senha:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=1, column=1,
                                                                                             sticky="ew")
            esenha_nova = Entry(frame2, show="*", width=20, font=fonte_padrao)
            esenha_nova.grid(row=1, column=2, sticky="w", padx=5, pady=10)
            esenha_nova.bind("<Return>", confirmar_bind)

            Label(frame2, text="Confirmar nova senha:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=2, column=1,
                                                                                             sticky="ew")
            esenha_confirma = Entry(frame2, show="*", width=20, font=fonte_padrao)
            esenha_confirma.grid(row=2, column=2, sticky="w", padx=5, pady=10)
            esenha_confirma.bind("<Return>", confirmar_bind)

            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(3, weight=1)

            bt1 = Button(frame3, text='Confirmar', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                         activeforeground="#FFFFFF", highlightthickness=0, width=10, relief=RIDGE, command=confirmar,
                         font=fonte_padrao, cursor="hand2")
            bt1.grid(row=0, column=1, pady=5, padx=5)
            bt2 = Button(frame3, text='Sair', width=10, relief=RIDGE, command=sair, font=fonte_padrao, cursor="hand2")
            bt2.grid(row=0, column=2, pady=5, padx=5)
            frame3.grid_columnconfigure(0, weight=1)
            frame3.grid_columnconfigure(3, weight=1)

            Label(frame4, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1, sticky="ew")
            frame4.grid_columnconfigure(0, weight=1)
            frame4.grid_columnconfigure(2, weight=1)
            '''root3.update()
            largura = frame0.winfo_width()
            altura = frame0.winfo_height()
            print(largura, altura)'''
            window_width = 340
            window_height = 247
            screen_width = root2.winfo_screenwidth()
            screen_height = root2.winfo_screenheight()
            x_cordinate = int((screen_width / 2) - (window_width / 2))
            y_cordinate = int((screen_height / 2) - (window_height / 2))
            root3.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
            #root3.resizable(0, 0)
            root3.iconbitmap('imagens\\ico.ico')
            root3.title(titulo_todos)
        def relatorio():
            root3 = Toplevel()
            root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
            root3.unbind_class("Button", "<Key-space>")
            root3.focus_force()
            root3.grab_set()
            def exportar():
                nome_coluna = str("")
                coluna = clique_colunas.get()
                filtro = ent_filtro.get()
                if filtro == "" and coluna != "":
                    messagebox.showwarning('Atenção:', 'O campo "Filtro" está vazio.', parent=root3)
                elif filtro != "" and coluna == "":
                    messagebox.showwarning('Atenção:', 'O campo "Coluna" está vazio.', parent=root3)
                else:
                    if coluna == "":
                        try:
                            cursor.execute(
                                "SELECT id_chamado, solicitante, data_abertura, tipo, titulo, ocorrencia, REPLACE(REPLACE(descricao_problema,CHAR(13),''),CHAR(10),''), nome_maquina, ramal, setor, data_atendimento, id_analista, status, prioridade, REPLACE(REPLACE(resolucao,CHAR(13),''),CHAR(10),''), data_encerramento, interacao FROM dbo.chamados ")
                            full_path = filedialog.asksaveasfilename(title="Exportar...",
                                                                     initialfile='Relatório HelpDesk (Completo)',
                                                                     filetypes=[('Excel', '.xlsx'),
                                                                                ('all files', '.*')],
                                                                     defaultextension='.xlsx')

                            workbook = xlsxwriter.Workbook(full_path)
                            worksheet = workbook.add_worksheet()
                            cell_format_cabecalho = workbook.add_format({'bold': True, 'font_color': '#1d366c'})
                            cell_format_cabecalho.set_align('center')

                            # Start from the first cell. Rows and columns are zero indexed.
                            row = 0
                            col = 0

                            worksheet.write(row, 0, 'Nº Chamado', cell_format_cabecalho)
                            worksheet.write(row, 1, 'Solicitante', cell_format_cabecalho)
                            worksheet.write(row, 2, 'Data de Abertura', cell_format_cabecalho)
                            worksheet.write(row, 3, 'Tipo', cell_format_cabecalho)
                            worksheet.write(row, 4, 'Título', cell_format_cabecalho)
                            worksheet.write(row, 5, 'Ocorrência', cell_format_cabecalho)
                            worksheet.write(row, 6, 'Descrição do Problema', cell_format_cabecalho)
                            worksheet.write(row, 7, 'Nome da Máquina', cell_format_cabecalho)
                            worksheet.write(row, 8, 'Ramal', cell_format_cabecalho)
                            worksheet.write(row, 9, 'Setor', cell_format_cabecalho)
                            worksheet.write(row, 10, 'Data de Atendimento', cell_format_cabecalho)
                            worksheet.write(row, 11, 'Analista', cell_format_cabecalho)
                            worksheet.write(row, 12, 'Status', cell_format_cabecalho)
                            worksheet.write(row, 13, 'Prioridade', cell_format_cabecalho)
                            worksheet.write(row, 14, 'Resolução', cell_format_cabecalho)
                            worksheet.write(row, 15, 'Data de Encerramento', cell_format_cabecalho)
                            worksheet.write(row, 16, 'Interação', cell_format_cabecalho)
                            row = 1

                            cell_format = workbook.add_format({'font_color': '#2c2c2c'})
                            cell_format.set_align('center')

                            # Iterate over the data and write it out row by row.
                            for teste in (cursor):
                                worksheet.write(row, col, teste[0], cell_format)
                                worksheet.write(row, col + 1, teste[1], cell_format)
                                worksheet.write(row, col + 2, teste[2], cell_format)
                                worksheet.write(row, col + 3, teste[3], cell_format)
                                worksheet.write(row, col + 4, teste[4], cell_format)
                                worksheet.write(row, col + 5, teste[5], cell_format)
                                worksheet.write(row, col + 6, teste[6], cell_format)
                                worksheet.write(row, col + 7, teste[7], cell_format)
                                worksheet.write(row, col + 8, teste[8], cell_format)
                                worksheet.write(row, col + 9, teste[9], cell_format)
                                worksheet.write(row, col + 10, teste[10], cell_format)
                                worksheet.write(row, col + 11, teste[11], cell_format)
                                worksheet.write(row, col + 12, teste[12], cell_format)
                                worksheet.write(row, col + 13, teste[13], cell_format)
                                worksheet.write(row, col + 14, teste[14], cell_format)
                                worksheet.write(row, col + 15, teste[15], cell_format)
                                worksheet.write(row, col + 16, teste[16], cell_format)
                                row += 1
                            workbook.close()
                        except:
                            messagebox.showwarning('Erro:', 'Erro ao exportar o relatório', parent=root3)
                            return False
                        messagebox.showinfo('Relatório:', 'Relatório exportado com sucesso.', parent=root3)
                        root3.destroy()
                        root2.destroy()
                    else:
                        if coluna == "Analista":
                            nome_coluna = "id_analista"
                        elif coluna == "Data de Abertura":
                            nome_coluna = "data_abertura"
                        elif coluna == "Data de Encerramento":
                            nome_coluna = "data_encerramento"
                        elif coluna == "Ocorrência":
                            nome_coluna = "ocorrencia"
                        elif coluna == "Prioridade":
                            nome_coluna = "prioridade"
                        elif coluna == "Setor":
                            nome_coluna = "setor"
                        elif coluna == "Solicitante":
                            nome_coluna = "solicitante"
                        elif coluna == "Status":
                            nome_coluna = "status"
                        query = "SELECT id_chamado, solicitante, data_abertura, tipo, titulo, ocorrencia, REPLACE(REPLACE(descricao_problema,CHAR(13),''),CHAR(10),''), nome_maquina, ramal, setor, data_atendimento, id_analista, status, prioridade, REPLACE(REPLACE(resolucao,CHAR(13),''),CHAR(10),''), data_encerramento, REPLACE(REPLACE(interacao,CHAR(13),''),CHAR(10),'') FROM dbo.chamados WHERE " + nome_coluna + " LIKE '%" + filtro + "%'"
                        cursor.execute(query)
                        busca = cursor.fetchone()
                        if busca is None:
                            messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root3)
                        else:
                            try:
                                full_path = filedialog.asksaveasfilename(title="Exportar...",
                                                                         initialfile='Relatório HelpDesk - Coluna(' + str(
                                                                             coluna) + ') Filtro(' + str(filtro) + ')',
                                                                         filetypes=[('Excel', '.xlsx'),
                                                                                    ('all files', '.*')],
                                                                         defaultextension='.xlsx')
                                workbook = xlsxwriter.Workbook(full_path)
                                worksheet = workbook.add_worksheet()
                                cell_format_cabecalho = workbook.add_format({'bold': True, 'font_color': '#1d366c'})
                                cell_format_cabecalho.set_align('center')

                                # Start from the first cell. Rows and columns are zero indexed.
                                row = 0
                                col = 0

                                worksheet.write(row, 0, 'Nº Chamado', cell_format_cabecalho)
                                worksheet.write(row, 1, 'Solicitante', cell_format_cabecalho)
                                worksheet.write(row, 2, 'Data de Abertura', cell_format_cabecalho)
                                worksheet.write(row, 3, 'Tipo', cell_format_cabecalho)
                                worksheet.write(row, 4, 'Título', cell_format_cabecalho)
                                worksheet.write(row, 5, 'Ocorrência', cell_format_cabecalho)
                                worksheet.write(row, 6, 'Descrição do Problema', cell_format_cabecalho)
                                worksheet.write(row, 7, 'Nome da Máquina', cell_format_cabecalho)
                                worksheet.write(row, 8, 'Ramal', cell_format_cabecalho)
                                worksheet.write(row, 9, 'Setor', cell_format_cabecalho)
                                worksheet.write(row, 10, 'Data de Atendimento', cell_format_cabecalho)
                                worksheet.write(row, 11, 'Analista', cell_format_cabecalho)
                                worksheet.write(row, 12, 'Status', cell_format_cabecalho)
                                worksheet.write(row, 13, 'Prioridade', cell_format_cabecalho)
                                worksheet.write(row, 14, 'Resolução', cell_format_cabecalho)
                                worksheet.write(row, 15, 'Data de Encerramento', cell_format_cabecalho)
                                worksheet.write(row, 16, 'Interação', cell_format_cabecalho)
                                row = 1

                                cell_format = workbook.add_format({'font_color': '#2c2c2c'})
                                cell_format.set_align('center')

                                # Iterate over the data and write it out row by row.
                                for teste in (cursor):
                                    worksheet.write(row, col, teste[0], cell_format)
                                    worksheet.write(row, col + 1, teste[1], cell_format)
                                    worksheet.write(row, col + 2, teste[2], cell_format)
                                    worksheet.write(row, col + 3, teste[3], cell_format)
                                    worksheet.write(row, col + 4, teste[4], cell_format)
                                    worksheet.write(row, col + 5, teste[5], cell_format)
                                    worksheet.write(row, col + 6, teste[6], cell_format)
                                    worksheet.write(row, col + 7, teste[7], cell_format)
                                    worksheet.write(row, col + 8, teste[8], cell_format)
                                    worksheet.write(row, col + 9, teste[9], cell_format)
                                    worksheet.write(row, col + 10, teste[10], cell_format)
                                    worksheet.write(row, col + 11, teste[11], cell_format)
                                    worksheet.write(row, col + 12, teste[12], cell_format)
                                    worksheet.write(row, col + 13, teste[13], cell_format)
                                    worksheet.write(row, col + 14, teste[14], cell_format)
                                    worksheet.write(row, col + 15, teste[15], cell_format)
                                    worksheet.write(row, col + 16, teste[16], cell_format)
                                    row += 1
                                workbook.close()
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao exportar o relatório', parent=root3)
                                return False

                            messagebox.showinfo('Relatório:', 'Relatório exportado com sucesso.', parent=root3)
                            root3.destroy()
                            root2.destroy()
            def exportar_bind(event):
                nome_coluna = str("")
                coluna = clique_colunas.get()
                filtro = ent_filtro.get()
                if filtro == "" and coluna != "":
                    messagebox.showwarning('Atenção:', 'O campo "Filtro" está vazio.', parent=root3)
                elif filtro != "" and coluna == "":
                    messagebox.showwarning('Atenção:', 'O campo "Coluna" está vazio.', parent=root3)
                else:
                    if coluna == "":
                        try:
                            cursor.execute(
                                "SELECT id_chamado, solicitante, data_abertura, tipo, titulo, ocorrencia, REPLACE(REPLACE(descricao_problema,CHAR(13),''),CHAR(10),''), nome_maquina, ramal, setor, data_atendimento, id_analista, status, prioridade, REPLACE(REPLACE(resolucao,CHAR(13),''),CHAR(10),''), data_encerramento   FROM dbo.chamados ")
                            full_path = filedialog.asksaveasfilename(title="Exportar...",
                                                                     initialfile='Relatório HelpDesk (Completo)',
                                                                     filetypes=[('Excel(CSV)', '.csv'),
                                                                                ('all files', '.*')],
                                                                     defaultextension='.csv')
                            with open(full_path, 'a') as outcsv:
                                writer = csv.writer(outcsv, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL,
                                                    lineterminator='\n')
                                writer.writerow(
                                    ['Nº Chamado', 'Solicitante', 'Data de Abertura', 'Tipo', 'Título', 'Ocorrência',
                                     'Descrição do Problema', 'Nome da Máquina', 'Ramal', 'Setor',
                                     'Data de Atendimento',
                                     'Analista', 'Status', 'Prioridade', 'Resolução', 'Data de Encerramento'])
                                for item in cursor:
                                    writer.writerow(
                                        [item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7],
                                         item[8],
                                         item[9], item[10], item[11], item[12], item[13], item[14], item[15]])
                            messagebox.showinfo('Relatório:', 'Relatório exportado com sucesso.', parent=root3)
                            root3.destroy()
                            root2.destroy()
                        except:
                            messagebox.showwarning('Erro:', 'Erro ao exportar o relatório', parent=root3)
                    else:
                        if coluna == "Analista":
                            nome_coluna = "id_analista"
                        elif coluna == "Data de Abertura":
                            nome_coluna = "data_abertura"
                        elif coluna == "Data de Encerramento":
                            nome_coluna = "data_encerramento"
                        elif coluna == "Ocorrência":
                            nome_coluna = "ocorrencia"
                        elif coluna == "Prioridade":
                            nome_coluna = "prioridade"
                        elif coluna == "Setor":
                            nome_coluna = "setor"
                        elif coluna == "Solicitante":
                            nome_coluna = "solicitante"
                        elif coluna == "Status":
                            nome_coluna = "status"
                        query = "SELECT id_chamado, solicitante, data_abertura, tipo, titulo, ocorrencia, REPLACE(REPLACE(descricao_problema,CHAR(13),''),CHAR(10),''), nome_maquina, ramal, setor, data_atendimento, id_analista, status, prioridade, REPLACE(REPLACE(resolucao,CHAR(13),''),CHAR(10),''), data_encerramento FROM dbo.chamados WHERE " + nome_coluna + " LIKE '%" + filtro + "%'"
                        cursor.execute(query)
                        busca = cursor.fetchone()
                        if busca is None:
                            messagebox.showwarning('Erro:', 'Nenhum registro encontrado', parent=root3)
                        else:
                            try:
                                full_path = filedialog.asksaveasfilename(title="Exportar...",
                                                                         initialfile='Relatório HelpDesk - Coluna(' + str(
                                                                             coluna) + ') Filtro(' + str(filtro) + ')',
                                                                         filetypes=[('Excel(CSV)', '.csv'),
                                                                                    ('all files', '.*')],
                                                                         defaultextension='.csv')
                                with open(full_path, 'a') as outcsv:
                                    writer = csv.writer(outcsv, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL,
                                                        lineterminator='\n')
                                    writer.writerow(
                                        ['Nº Chamado', 'Solicitante', 'Data de Abertura', 'Tipo', 'Título',
                                         'Ocorrência',
                                         'Descrição do Problema', 'Nome da Máquina', 'Ramal', 'Setor',
                                         'Data de Atendimento',
                                         'Analista', 'Status', 'Prioridade', 'Resolução', 'Data de Encerramento'])
                                    for item in cursor:
                                        writer.writerow(
                                            [item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7],
                                             item[8],
                                             item[9], item[10], item[11], item[12], item[13], item[14], item[15]])
                                messagebox.showinfo('Relatório:', 'Relatório exportado com sucesso.', parent=root3)
                                root3.destroy()
                                root2.destroy()

                            except:
                                messagebox.showwarning('Erro:', 'Erro ao exportar o relatório', parent=root3)
            def limpar_filtros():
                clique_colunas.set('')
                ent_filtro.delete(0,END)

            frame0 = Frame(root3, bg='#ffffff')
            frame0.grid(row=0, column=0, stick='nsew')
            root3.grid_rowconfigure(0, weight=1)
            root3.grid_columnconfigure(0, weight=1)
            frame1 = Frame(frame0, bg="#1d366c")
            frame1.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame2 = Frame(frame0, bg='#ffffff', pady=6)
            frame2.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame3 = Frame(frame0, bg='#1d366c')
            frame3.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame4 = Frame(frame0, bg='#ffffff', pady=10)
            frame4.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame5 = Frame(frame0, bg='#ffffff')
            frame5.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame6 = Frame(frame0, bg='#1d366c')
            frame6.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame7 = Frame(frame0, bg='#ffffff')
            frame7.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame8 = Frame(frame0, bg='#1d366c')
            frame8.pack(side=TOP, fill=X, expand=False, anchor='n')

            image_trocasenha2 = Image.open('imagens\\relatorio2.png')
            resize_trocasenha2 = image_trocasenha2.resize((35, 35))
            nova_image_trocasenha2 = ImageTk.PhotoImage(resize_trocasenha2)

            lbllogin = Label(frame1, image=nova_image_trocasenha2, text=" Relatório", compound="left", bg='#1d366c', fg='#FFFFFF', font=fonte_titulos)
            lbllogin.photo = nova_image_trocasenha2
            lbllogin.grid(row=0, column=1)
            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(2, weight=1)

            Label(frame2, text="Configurações:", font=fonte_titulos, bg='#ffffff', fg='#000000').grid(row=0, column=1, sticky="ew")
            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(2, weight=1)

            #frame3 linha horizontal.

            Label(frame4, text="Tabela:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=1, column=1,sticky="ew")
            clique_tabela = StringVar()
            drop_tabela = OptionMenu(frame4, clique_tabela, 'Chamados', 'Analistas', command=drop_selecao_busca)
            drop_tabela.config(font=fonte_padrao, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                              activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=19, cursor="hand2")
            drop_tabela.grid(row=1, column=2, sticky="w", padx=5, pady=10)
            clique_tabela.set('Chamados')
            drop_tabela.config(state=DISABLED, bg='#BDBDBD', fg='#FFFFFF', activebackground='#BDBDBD',
                            activeforeground="#BDBDBD", highlightthickness=0, relief=RIDGE, width=19)

            Label(frame4, text="Coluna:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=2, column=1,sticky="ew")
            clique_colunas = StringVar()
            drop_colunas = OptionMenu(frame4, clique_colunas, 'Analista', 'Data de Abertura', 'Data de Encerramento', 'Ocorrência', 'Prioridade', 'Setor', 'Solicitante', 'Status', command=drop_selecao_busca)
            drop_colunas.config(font=fonte_padrao, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                              activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=19, cursor="hand2")
            drop_colunas.grid(row=2, column=2, sticky="w", padx=5, pady=10)

            Label(frame4, text="Filtro:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=3, column=1,sticky="ew")
            ent_filtro = Entry(frame4, width=24, font=fonte_padrao, justify='center')
            ent_filtro.grid(row=3, column=2, sticky="w", padx=5, pady=10)
            ent_filtro.bind('<Return>', exportar_bind)

            frame4.grid_columnconfigure(0, weight=1)
            frame4.grid_columnconfigure(3, weight=1)

            Label(frame5, text="Obs: Para exportar o relatório completo, os campos devem estar vazios.", font=fonte_botao, bg='#ffffff', fg='#4f4f4f').grid(row=0, column=1, sticky="ew", pady=(0,8))
            frame5.grid_columnconfigure(0, weight=1)
            frame5.grid_columnconfigure(2, weight=1)
            #frame6 linha horizontal

            bt1 = Button(frame7, text='Limpar Filtros', width=12, relief=RIDGE, command=limpar_filtros, font=fonte_botao, cursor="hand2")
            bt1.grid(row=0, column=1, pady=5, padx=5)
            bt2 = Button(frame7, text='Exportar', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                         activeforeground="#FFFFFF", highlightthickness=0, width=12, relief=RIDGE, command=exportar,
                         font=fonte_botao, cursor="hand2")
            bt2.grid(row=0, column=2, pady=5, padx=5)
            frame7.grid_columnconfigure(0, weight=1)
            frame7.grid_columnconfigure(3, weight=1)


            Label(frame8, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1, sticky="ew")
            frame8.grid_columnconfigure(0, weight=1)
            frame8.grid_columnconfigure(2, weight=1)

            '''root3.update()
            largura = frame0.winfo_width()
            altura = frame0.winfo_height()
            print(largura, altura)'''
            window_width = 392
            window_height = 321
            screen_width = root2.winfo_screenwidth()
            screen_height = root2.winfo_screenheight()
            x_cordinate = int((screen_width / 2) - (window_width / 2))
            y_cordinate = int((screen_height / 2) - (window_height / 2))
            root3.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
            root3.resizable(0, 0)
            root3.iconbitmap('imagens\\ico.ico')
            root3.title(titulo_todos)
        def estoque():
            root3 = Toplevel()
            root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
            root3.unbind_class("Button", "<Key-space>")
            root3.focus_force()
            root3.grab_set()

            def home():
                #os.system('cmd /c "net use //192.168.1.19/hdgv /user:impressoras gv2K17ADM"')


                for widget in frame4.winfo_children():
                    widget.destroy()
                fr0 = Frame(frame4, bg='#ffffff')
                fr0.pack(side=TOP, fill=BOTH)

                fr1 = Frame(fr0, bg='#ffffff')
                fr1.pack(side=TOP, fill=X, expand=False, anchor='n')
                fr2 = Frame(fr0, bg='#1d366c')
                fr2.pack(side=TOP, fill=X, expand=True, anchor='n', pady=8)


                image_entrada = Image.open('imagens\\home.png')
                resize_entrada = image_entrada.resize((30, 30))
                nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                lbl = Label(fr1, text=" Home", image=nova_image_entrada, font=fonte_titulos, compound="left", bg="#ffffff", fg='#000000')
                lbl.grid(row=0, column=1)
                fr1.grid_columnconfigure(0, weight=1)
                fr1.grid_columnconfigure(2, weight=1)

                # fr2 linha horizontal

                style = ttk.Style()
                # style.theme_use('default')
                style.configure('Treeview',
                                background='#ffffff',
                                rowheight=24,
                                fieldbackground='#ffffff',
                                font=fonte_padrao)
                style.configure("Treeview.Heading",
                                foreground='#000000',
                                background="#ffffff",
                                font=fonte_padrao)
                style.map('Treeview', background=[('selected', '#1d366c')])

                tree_principal = ttk.Treeview(frame4, selectmode='none')
                vsb = ttk.Scrollbar(frame4, orient="vertical", command=tree_principal.yview)
                vsb.pack(side=RIGHT, fill='y')
                tree_principal.configure(yscrollcommand=vsb.set)
                vsbx = ttk.Scrollbar(frame4, orient="horizontal", command=tree_principal.xview)
                vsbx.pack(side=BOTTOM, fill='x')
                tree_principal.configure(xscrollcommand=vsbx.set)
                tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
                tree_principal["columns"] = ("1", "2", "3")
                tree_principal['show'] = 'headings'

                tree_principal.column("1", width=88, anchor='c')
                tree_principal.column("2", width=116, anchor='c')
                tree_principal.column("3", width=180, anchor='c')
                tree_principal.heading("1", text="Nome do Produto")
                tree_principal.heading("2", text="Quantidade em Estoque")
                tree_principal.heading("3", text="Estoque Mínimo")
                tree_principal.tag_configure('par', background='#A52A2A')
                tree_principal.tag_configure('impar', background='#ffffff')
                cursor.execute("SELECT * FROM dbo.produtos ORDER BY nome_produto ASC")
                for row in cursor:
                    if row[3] <= row[4]:
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                                  row[1], row[3], row[4]),
                                              tags=('par',))
                    else:
                        tree_principal.insert('', 'end', text=" ",
                                              values=(
                                                  row[1], row[3], row[4]),
                                              tags=('impar',))
                root3.mainloop()

            def entrada():
                def atualizar_lista():
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.produtos ORDER BY nome_produto ASC")
                    cont = 0
                    for row in cursor:
                        if cont % 2 == 0:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[3], row[4]),
                                                  tags=('par',))
                        else:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[3], row[4]),
                                                  tags=('impar',))
                        cont += 1
                def duplo_clique(event):
                    produto_select = tree_principal.focus()
                    nome_produto = tree_principal.item(produto_select, "values")[0]
                    entrynome_produto.config(state='normal')
                    entrynome_produto.delete(0, END)
                    entrynome_produto.insert(0, nome_produto)
                    entrynome_produto.config(state='disabled')
                def salvar():
                    if entrynome_produto.get() == "" or entrypedido.get() == "" or entryquanti.get() == "":
                        messagebox.showwarning('Erro:', 'Todos os campos são obrigatórios.', parent=root3)
                    else:
                        try:
                            n_pedido =int(entrypedido.get())
                        except:
                            messagebox.showwarning('Erro:', 'Somente números inteiros são permitidos.', parent=root3)
                            return False
                        try:
                            quantidade =int(entryquanti.get())
                        except:
                            messagebox.showwarning('Erro:', 'Somente números inteiros são permitidos.', parent=root3)
                            return False

                        nome_produto = entrynome_produto.get()
                        r = cursor.execute("SELECT * FROM dbo.produtos WHERE nome_produto=?", (nome_produto,))
                        result = r.fetchone()

                        soma_quantidade = int(result[3]) + quantidade

                        resposta = messagebox.askyesno('Atenção:',
                                                       f'Tem certeza de que deseja dar entrada neste produto ({nome_produto})?',
                                                       parent=root3)
                        if resposta == True:
                            try:
                                cursor.execute(
                                    "UPDATE helpdesk.dbo.produtos SET quantidade = ? WHERE nome_produto = ?",
                                    (soma_quantidade, nome_produto))
                                cursor.commit()
                            except:
                                messagebox.showerror('Erro:', 'Erro ao atualizar o banco de dados (Tabela:produtos).', parent=root3)
                                return False
                            try:
                                tupla = (nome_produto, n_pedido, quantidade,data,usuariologado)
                                cursor.execute("INSERT INTO dbo.entrada (produto,pedido,quantidade_entrada,data_entrada,analista ) values(?,?,?,?,?)", (tupla))
                                cursor.commit()
                                messagebox.showinfo('Sucesso:', 'Entrada realizada com sucesso.', parent=root3)
                                home()
                            except:
                                messagebox.showerror('Erro:', 'Erro ao atualizar o banco de dados (Tabela:entrada).', parent=root3)
                                return False

                for widget in frame4.winfo_children():
                    widget.destroy()
                fr0 = Frame(frame4, bg='#ffffff')
                fr0.pack(side=TOP, fill=BOTH)

                fr1 = Frame(fr0, bg='#ffffff')
                fr1.pack(side=TOP, fill=X, expand=False, anchor='n')
                fr2 = Frame(fr0, bg='#1d366c')
                fr2.pack(side=TOP, fill=X, expand=True, anchor='n', pady=8)
                fr3 = Frame(fr0, bg='#ffffff')
                fr3.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr4 = Frame(fr0, bg='#1d366c')
                fr4.pack(side=TOP, fill=BOTH, expand=True, anchor='n', pady=8)
                fr5 = Frame(fr0, bg='#ffffff')
                fr5.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr6 = Frame(fr0, bg='#ffffff')
                fr6.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr7 = Frame(fr0, bg='#1d366c')
                fr7.pack(side=BOTTOM, fill=X, expand=True, anchor='n', pady=8)

                image_entrada = Image.open('imagens\\entrada.png')
                resize_entrada = image_entrada.resize((30, 30))
                nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                lbl = Label(fr1, text=" Entrada de Produtos", image=nova_image_entrada, font=fonte_titulos, compound="left", bg="#ffffff", fg='#000000')
                lbl.grid(row=0, column=1)
                fr1.grid_columnconfigure(0, weight=1)
                fr1.grid_columnconfigure(2, weight=1)

                # fr2 linha horizontal

                lblnome_produto = Label(fr3, text="Produto:", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblnome_produto.grid(row=1, column=1, pady=4)
                entrynome_produto = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entrynome_produto.grid(row=1, column=2, pady=4)
                entrynome_produto.config(state='disabled')

                lblpedido = Label(fr3, text="Nº Pedido:", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblpedido.grid(row=2, column=1, pady=4)
                entrypedido = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entrypedido.grid(row=2, column=2, pady=4)

                lblquanti = Label(fr3, text="Quantidade (Entrada):", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblquanti.grid(row=3, column=1, pady=4)
                entryquanti = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entryquanti.grid(row=3, column=2, pady=4)



                fr3.grid_columnconfigure(0, weight=1)
                fr3.grid_columnconfigure(3, weight=1)

                # fr4 linha

                btn_cadastro = Button(fr5, text='Adicionar', bg='#1d366c', fg='#FFFFFF',
                                      activebackground='#1d366c',
                                      activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                                      command=salvar, font=fonte_botao, cursor="hand2")
                btn_cadastro.grid(row=0, column=1, pady=4)

                fr5.grid_columnconfigure(0, weight=1)
                fr5.grid_columnconfigure(2, weight=1)

                style = ttk.Style()
                # style.theme_use('default')
                style.configure('Treeview',
                                background='#ffffff',
                                rowheight=24,
                                fieldbackground='#ffffff',
                                font=fonte_padrao)
                style.configure("Treeview.Heading",
                                foreground='#000000',
                                background="#ffffff",
                                font=fonte_padrao)
                style.map('Treeview', background=[('selected', '#1d366c')])

                tree_principal = ttk.Treeview(frame4, selectmode='browse')
                vsb = ttk.Scrollbar(frame4, orient="vertical", command=tree_principal.yview)
                vsb.pack(side=RIGHT, fill='y')
                tree_principal.configure(yscrollcommand=vsb.set)
                vsbx = ttk.Scrollbar(frame4, orient="horizontal", command=tree_principal.xview)
                vsbx.pack(side=BOTTOM, fill='x')
                tree_principal.configure(xscrollcommand=vsbx.set)
                tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
                tree_principal["columns"] = ("1", "2", "3")
                tree_principal['show'] = 'headings'
                tree_principal.column("1", width=88, anchor='c')
                tree_principal.column("2", width=116, anchor='c')
                tree_principal.column("3", width=180, anchor='c')
                tree_principal.heading("1", text="Nome do Produto")
                tree_principal.heading("2", text="Quantidade em Estoque")
                tree_principal.heading("3", text="Estoque Mínimo")
                tree_principal.tag_configure('par', background='#e9e9e9')
                tree_principal.tag_configure('impar', background='#ffffff')
                tree_principal.bind("<Double-1>", duplo_clique)
                atualizar_lista()
                root3.mainloop()

            def saida():
                def atualizar_lista():
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.produtos ORDER BY nome_produto ASC")
                    cont = 0
                    for row in cursor:
                        if cont % 2 == 0:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[3], row[4]),
                                                  tags=('par',))
                        else:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[3], row[4]),
                                                  tags=('impar',))
                        cont += 1
                def duplo_clique(event):
                    produto_select = tree_principal.focus()
                    nome_produto = tree_principal.item(produto_select, "values")[0]
                    entrynome_produto.config(state='normal')
                    entrynome_produto.delete(0, END)
                    entrynome_produto.insert(0, nome_produto)
                    entrynome_produto.config(state='disabled')
                def salvar():
                    if entrynome_produto.get() == "" or entrysolic.get() == "" or clique_estoque_setor.get() == "" or entryquanti.get() =="":
                        messagebox.showwarning('Erro:', 'Todos os campos são obrigatórios.', parent=root3)
                    else:
                        try:
                            quantidade =int(entryquanti.get())
                        except:
                            messagebox.showwarning('Erro:', 'Somente números inteiros são permitidos.', parent=root3)
                            return False

                        nome_produto = entrynome_produto.get()
                        setor = clique_estoque_setor.get()
                        solicitante = entrysolic.get().upper()

                        r = cursor.execute("SELECT * FROM dbo.produtos WHERE nome_produto=?", (nome_produto,))
                        result = r.fetchone()

                        sub_quantidade = int(result[3]) - quantidade
                        print(sub_quantidade)
                        resposta = messagebox.askyesno('Atenção:',
                                                       f'Tem certeza de que deseja dar saída neste produto ({nome_produto})?',
                                                       parent=root3)
                        if resposta == True:
                            try:
                                cursor.execute(
                                    "UPDATE helpdesk.dbo.produtos SET quantidade = ? WHERE nome_produto = ?",
                                    (sub_quantidade, nome_produto))
                                cursor.commit()
                            except:
                                messagebox.showerror('Erro:', 'Erro ao atualizar o banco de dados (Tabela:produtos).', parent=root3)
                                return False
                            try:
                                tupla = (nome_produto, setor, solicitante, quantidade, data, usuariologado)
                                cursor.execute("INSERT INTO dbo.saida(produto, setor, solicitante, quantidade_saida, data_saida, analista ) values(?,?,?,?,?,?)", (tupla))
                                cursor.commit()
                                messagebox.showinfo('Sucesso:', 'Saída realizada com sucesso.', parent=root3)
                                home()
                            except:
                                messagebox.showerror('Erro:', 'Erro ao atualizar o banco de dados (Tabela:saída).', parent=root3)
                                return False


                for widget in frame4.winfo_children():
                    widget.destroy()
                fr0 = Frame(frame4, bg='#ffffff')
                fr0.pack(side=TOP, fill=BOTH)

                fr1 = Frame(fr0, bg='#ffffff')
                fr1.pack(side=TOP, fill=X, expand=False, anchor='n')
                fr2 = Frame(fr0, bg='#1d366c')
                fr2.pack(side=TOP, fill=X, expand=True, anchor='n', pady=8)
                fr3 = Frame(fr0, bg='#ffffff')
                fr3.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr4 = Frame(fr0, bg='#1d366c')
                fr4.pack(side=TOP, fill=BOTH, expand=True, anchor='n', pady=8)
                fr5 = Frame(fr0, bg='#ffffff')
                fr5.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr6 = Frame(fr0, bg='#ffffff')
                fr6.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr7 = Frame(fr0, bg='#1d366c')
                fr7.pack(side=BOTTOM, fill=X, expand=True, anchor='n', pady=8)

                image_entrada = Image.open('imagens\\saida.png')
                resize_entrada = image_entrada.resize((30, 30))
                nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                lbl = Label(fr1, text=" Saída de Produtos", image=nova_image_entrada, font=fonte_titulos, compound="left", bg="#ffffff", fg='#000000')
                lbl.grid(row=0, column=1)
                fr1.grid_columnconfigure(0, weight=1)
                fr1.grid_columnconfigure(2, weight=1)

                # fr2 linha horizontal

                lblnome_produto = Label(fr3, text="Produto:", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblnome_produto.grid(row=1, column=1, pady=4)
                entrynome_produto = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entrynome_produto.grid(row=1, column=2, pady=4)
                entrynome_produto.config(state='disabled')

                lblsolic = Label(fr3, text="Solicitante:", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblsolic.grid(row=2, column=1, pady=4)
                entrysolic = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entrysolic.grid(row=2, column=2, pady=4)

                lblsetor = Label(fr3, text="Setor:", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblsetor.grid(row=3, column=1, pady=4)
                OptionList = [
                    "Aciaria",
                    "Almoxarifado",
                    "Ambulatório",
                    "Balança",
                    "Comercial",
                    "Compras",
                    "Contabilidade",
                    "Custos",
                    "EHS",
                    "Elétrica",
                    "Engenharia",
                    "Faturamento",
                    "Financeiro",
                    "Fiscal",
                    "Lab Inspeção",
                    "Lab Mecânico",
                    "Lab Químico",
                    "Laminação",
                    "Logística",
                    "Oficina de Cilindros",
                    "Oficina Mecânica",
                    "Pátio de Sucata",
                    "PCP",
                    "Planta D´agua",
                    "Planta de Escória",
                    "Portaria",
                    "Qualidade",
                    "Refratários",
                    "RH",
                    "Subestação",
                    "TI",
                    "Utilidades"
                ]
                clique_estoque_setor = StringVar()
                drop_setor = OptionMenu(fr3, clique_estoque_setor, *OptionList)
                drop_setor.config(bg='#ffffff', fg='#000000', activebackground='#dcdcdc', activeforeground="#000000",
                                  highlightthickness=0, relief=RIDGE, width=41, cursor="hand2")
                drop_setor.grid(row=3, column=2, pady=4)

                lblquanti = Label(fr3, text="Quantidade (Saída):", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblquanti.grid(row=4, column=1, pady=4)
                entryquanti = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entryquanti.grid(row=4, column=2, pady=4)

                fr3.grid_columnconfigure(0, weight=1)
                fr3.grid_columnconfigure(3, weight=1)

                # fr4 linha

                btn_cadastro = Button(fr5, text='Saída', bg='#1d366c', fg='#FFFFFF',
                                      activebackground='#1d366c',
                                      activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                                      command=salvar, font=fonte_botao, cursor="hand2")
                btn_cadastro.grid(row=0, column=1, pady=4)

                fr5.grid_columnconfigure(0, weight=1)
                fr5.grid_columnconfigure(2, weight=1)

                style = ttk.Style()
                # style.theme_use('default')
                style.configure('Treeview',
                                background='#ffffff',
                                rowheight=24,
                                fieldbackground='#ffffff',
                                font=fonte_padrao)
                style.configure("Treeview.Heading",
                                foreground='#000000',
                                background="#ffffff",
                                font=fonte_padrao)
                style.map('Treeview', background=[('selected', '#1d366c')])

                tree_principal = ttk.Treeview(frame4, selectmode='browse')
                vsb = ttk.Scrollbar(frame4, orient="vertical", command=tree_principal.yview)
                vsb.pack(side=RIGHT, fill='y')
                tree_principal.configure(yscrollcommand=vsb.set)
                vsbx = ttk.Scrollbar(frame4, orient="horizontal", command=tree_principal.xview)
                vsbx.pack(side=BOTTOM, fill='x')
                tree_principal.configure(xscrollcommand=vsbx.set)
                tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
                tree_principal["columns"] = ("1", "2", "3")
                tree_principal['show'] = 'headings'
                tree_principal.column("1", width=88, anchor='c')
                tree_principal.column("2", width=116, anchor='c')
                tree_principal.column("3", width=180, anchor='c')
                tree_principal.heading("1", text="Nome do Produto")
                tree_principal.heading("2", text="Quantidade em Estoque")
                tree_principal.heading("3", text="Estoque Mínimo")
                tree_principal.tag_configure('par', background='#e9e9e9')
                tree_principal.tag_configure('impar', background='#ffffff')
                tree_principal.bind("<Double-1>", duplo_clique)
                atualizar_lista()
                root3.mainloop()

            def historico():
                def hist_entrada():
                    for widget in fr5.winfo_children():
                        widget.destroy()
                    for widget in fr7.winfo_children():
                        widget.destroy()

                    image_entrada = Image.open('imagens\\entrada.png')
                    resize_entrada = image_entrada.resize((30, 30))
                    nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                    lbl = Label(fr5, text=" Histórico de Entrada", image=nova_image_entrada, font=fonte_titulos, compound="left",
                                bg="#ffffff", fg='#000000')
                    lbl.grid(row=0, column=1)
                    fr5.grid_columnconfigure(0, weight=1)
                    fr5.grid_columnconfigure(2, weight=1)

                    style = ttk.Style()
                    # style.theme_use('default')
                    style.configure('Treeview',
                                    background='#ffffff',
                                    rowheight=24,
                                    fieldbackground='#ffffff',
                                    font=fonte_padrao)
                    style.configure("Treeview.Heading",
                                    foreground='#000000',
                                    background="#ffffff",
                                    font=fonte_padrao)
                    style.map('Treeview', background=[('selected', '#1d366c')])

                    tree_principal = ttk.Treeview(fr7, selectmode='none')
                    vsb = ttk.Scrollbar(fr7, orient="vertical", command=tree_principal.yview)
                    vsb.pack(side=RIGHT, fill='y')
                    tree_principal.configure(yscrollcommand=vsb.set)
                    vsbx = ttk.Scrollbar(fr7, orient="horizontal", command=tree_principal.xview)
                    vsbx.pack(side=BOTTOM, fill='x')
                    tree_principal.configure(xscrollcommand=vsbx.set)
                    tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
                    tree_principal["columns"] = ("1", "2", "3", "4", "5")
                    tree_principal['show'] = 'headings'
                    tree_principal.column("1", width=100, anchor='c')
                    tree_principal.column("2", width=100, anchor='c')
                    tree_principal.column("3", width=100, anchor='c')
                    tree_principal.column("4", width=100, anchor='c')
                    tree_principal.column("5", width=100, anchor='c')
                    tree_principal.heading("1", text="Nome do Produto")
                    tree_principal.heading("2", text="Nº do Pedido")
                    tree_principal.heading("3", text="Quantidade de Entrada")
                    tree_principal.heading("4", text="Data de Entrada")
                    tree_principal.heading("5", text="Analista Responsável")
                    tree_principal.tag_configure('par', background='#e9e9e9')
                    tree_principal.tag_configure('impar', background='#ffffff')


                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.entrada ORDER BY produto ASC")
                    cont = 0
                    for row in cursor:
                        if cont % 2 == 0:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(row[1], row[2], row[3], row[4], row[5]), tags=('par',))
                        else:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[2], row[3], row[4], row[5]),
                                                  tags=('impar',))
                        cont += 1


                    root3.mainloop()

                def hist_saida():
                    for widget in fr5.winfo_children():
                        widget.destroy()
                    for widget in fr7.winfo_children():
                        widget.destroy()
                    image_entrada = Image.open('imagens\\saida.png')
                    resize_entrada = image_entrada.resize((30, 30))
                    nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                    lbl = Label(fr5, text=" Histórico de Saída", image=nova_image_entrada, font=fonte_titulos, compound="left",
                                bg="#ffffff", fg='#000000')
                    lbl.grid(row=0, column=1)
                    fr5.grid_columnconfigure(0, weight=1)
                    fr5.grid_columnconfigure(2, weight=1)

                    style = ttk.Style()
                    # style.theme_use('default')
                    style.configure('Treeview',
                                    background='#ffffff',
                                    rowheight=24,
                                    fieldbackground='#ffffff',
                                    font=fonte_padrao)
                    style.configure("Treeview.Heading",
                                    foreground='#000000',
                                    background="#ffffff",
                                    font=fonte_padrao)
                    style.map('Treeview', background=[('selected', '#1d366c')])

                    tree_principal = ttk.Treeview(fr7, selectmode='none')
                    vsb = ttk.Scrollbar(fr7, orient="vertical", command=tree_principal.yview)
                    vsb.pack(side=RIGHT, fill='y')
                    tree_principal.configure(yscrollcommand=vsb.set)
                    vsbx = ttk.Scrollbar(fr7, orient="horizontal", command=tree_principal.xview)
                    vsbx.pack(side=BOTTOM, fill='x')
                    tree_principal.configure(xscrollcommand=vsbx.set)
                    tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
                    tree_principal["columns"] = ("1", "2", "3", "4", "5", "6")
                    tree_principal['show'] = 'headings'
                    tree_principal.column("1", width=100, anchor='c')
                    tree_principal.column("2", width=100, anchor='c')
                    tree_principal.column("3", width=100, anchor='c')
                    tree_principal.column("4", width=100, anchor='c')
                    tree_principal.column("5", width=100, anchor='c')
                    tree_principal.column("6", width=100, anchor='c')
                    tree_principal.heading("1", text="Nome do Produto")
                    tree_principal.heading("2", text="Setor")
                    tree_principal.heading("3", text="Solicitante")
                    tree_principal.heading("4", text="Quantidade de Saída")
                    tree_principal.heading("5", text="Data de Saída")
                    tree_principal.heading("6", text="Analista Responsável")
                    tree_principal.tag_configure('par', background='#e9e9e9')
                    tree_principal.tag_configure('impar', background='#ffffff')

                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.saida ORDER BY produto ASC")
                    cont = 0
                    for row in cursor:
                        print
                        if cont % 2 == 0:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(row[1], row[2], row[3], row[4], row[5], row[6]), tags=('par',))
                        else:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[2], row[3], row[4], row[5], row[6]),
                                                  tags=('impar',))
                        cont += 1


                    root3.mainloop()

                for widget in frame4.winfo_children():
                    widget.destroy()
                fr0 = Frame(frame4, bg='#ffffff')
                fr0.pack(side=TOP, fill=BOTH, expand=True)

                fr1 = Frame(fr0, bg='#ffffff')
                fr1.pack(side=TOP, fill=X, expand=False, anchor='n')
                fr2 = Frame(fr0, bg='#1d366c')
                fr2.pack(side=TOP, fill=X, expand=False, anchor='n', pady=8)
                fr3 = Frame(fr0, bg='#ffffff')
                fr3.pack(side=TOP, fill=X, expand=False, anchor='n')
                fr4 = Frame(fr0, bg='#1d366c')
                fr4.pack(side=TOP, fill=X, expand=False, anchor='n', pady=8)
                fr5 = Frame(fr0, bg='#ffffff')
                fr5.pack(side=TOP, fill=X, expand=False, anchor='n')
                fr6 = Frame(fr0, bg='#1d366c')
                fr6.pack(side=TOP, fill=X, expand=False, anchor='n', pady=8)
                fr7 = Frame(fr0, bg='#ffffff')
                fr7.pack(side=TOP, fill=BOTH, expand=True, anchor='n')


                image_entrada = Image.open('imagens\\historico.png')
                resize_entrada = image_entrada.resize((30, 30))
                nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                lbl = Label(fr1, text=" Histórico", image=nova_image_entrada, font=fonte_titulos, compound="left", bg="#ffffff", fg='#000000')
                lbl.grid(row=0, column=1)
                fr1.grid_columnconfigure(0, weight=1)
                fr1.grid_columnconfigure(2, weight=1)

                # fr2 linha horizontal


                btn_entrada = Button(fr3, text='Histórico de Entrada', bg='#dcdcdc', fg='#000000',
                                      activebackground='#1d366c',
                                      activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                                      command=hist_entrada, font=fonte_botao, cursor="hand2")
                btn_entrada.grid(row=0, column=1, pady=4, padx=4)

                btn_saida = Button(fr3, text='Histórico de Saída', bg='#dcdcdc', fg='#000000',
                                      activebackground='#1d366c',
                                      activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                                      command=hist_saida, font=fonte_botao, cursor="hand2")
                btn_saida.grid(row=0, column=2, pady=4, padx=4)

                fr3.grid_columnconfigure(0, weight=1)
                fr3.grid_columnconfigure(3, weight=1)

                # fr4 linha

                hist_entrada()

                root3.mainloop()

            def cadastro():
                def atualizar_lista():
                    tree_principal.delete(*tree_principal.get_children())
                    cursor.execute("SELECT * FROM dbo.produtos ORDER BY nome_produto ASC")
                    cont = 0
                    for row in cursor:
                        if cont % 2 == 0:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[2], row[3], row[4]),
                                                  tags=('par',))
                        else:
                            tree_principal.insert('', 'end', text=" ",
                                                  values=(
                                                  row[1], row[2], row[3], row[4]),
                                                  tags=('impar',))
                        cont += 1
                def conversor(filename):
                    nome_arquivo_origem = os.path.basename(filename).upper()
                    verifica = cursor.execute("SELECT * FROM dbo.produtos WHERE nome_imagem=?", (nome_arquivo_origem,))
                    compara = verifica.fetchone()
                    if compara == None:
                        return nome_arquivo_origem
                    else:
                        messagebox.showwarning('Erro:', f'Imagem já cadastrada ({nome_arquivo_origem}).', parent=root3)
                def upload_imagem():
                    global caminho_origem
                    caminho_origem = filedialog.askopenfilename(initialdir="os.path.expanduser(default_dir)",
                                                       title="Escolha um Arquivo",filetypes=([("PNG", "*.png")]))
                    global nome_imagem_verificada
                    nome_imagem_verificada = conversor(caminho_origem)
                def salvar():
                    try:
                        estoque_min = int(entryest_minimo.get())
                    except:
                        messagebox.showwarning('Erro:', 'Somente números inteiros são permitidos.', parent=root3)
                        return False
                    try:
                        nome_arquivo_origem = nome_imagem_verificada
                    except:
                        messagebox.showwarning('Erro:', 'Escolha uma imagem para o produto.', parent=root3)
                        return False
                    nome_produto = entrynome_produto.get().upper()
                    if nome_produto == "":
                        messagebox.showwarning('Erro:', 'Campo "Nome do Produto" vazio.', parent=root3)
                    else:
                        ext = os.path.splitext(nome_arquivo_origem)
                        novo_nome_arquivo_origem = str(nome_produto) + str(ext[1])
                        print(novo_nome_arquivo_origem)
                        r = cursor.execute("SELECT * FROM dbo.produtos WHERE nome_produto=?", (nome_produto,))
                        compara_nome = r.fetchone()
                        if compara_nome == None:
                            destino_local = 'imagens/estoque/' + novo_nome_arquivo_origem
                            destino_rede = r'\\192.168.1.19/hdGV2/imagens/estoque/' + novo_nome_arquivo_origem
                            try:
                                shutil.copy(caminho_origem, destino_local)
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao copiar o arquivo para a máquina local.',
                                                       parent=root3)
                            try:
                                shutil.copy(caminho_origem, destino_rede)
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao copiar o arquivo para o servidor.',
                                                       parent=root3)
                                return False
                            try:
                                tupla = (nome_produto, novo_nome_arquivo_origem, 0, estoque_min)
                                cursor.execute("INSERT INTO dbo.produtos (nome_produto, nome_imagem, quantidade, estq_minimo) values(?,?,?,?)",
                                               (tupla))
                                cursor.commit()
                                messagebox.showinfo('Sucesso:', 'Produto cadastrado com sucesso.', parent=root3)
                                home()
                            except:
                                messagebox.showwarning('Erro:', 'Erro ao gravar as informações no banco de dados.',
                                                       parent=root3)

                        else:
                            messagebox.showinfo('Atenção:', 'Já existe um "Produto" com o mesmo nome cadastrado.', parent=root3)

                for widget in frame4.winfo_children():
                    widget.destroy()
                fr0 = Frame(frame4, bg='#ffffff')
                fr0.pack(side=TOP, fill=BOTH)

                fr1 = Frame(fr0, bg='#ffffff')
                fr1.pack(side=TOP, fill=X, expand=False, anchor='n')
                fr2 = Frame(fr0, bg='#1d366c')
                fr2.pack(side=TOP, fill=X, expand=True, anchor='n', pady=8)
                fr3 = Frame(fr0, bg='#ffffff')
                fr3.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr4 = Frame(fr0, bg='#1d366c')
                fr4.pack(side=TOP, fill=BOTH, expand=True, anchor='n', pady=8)
                fr5 = Frame(fr0, bg='#ffffff')
                fr5.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr6 = Frame(fr0, bg='#ffffff')
                fr6.pack(side=TOP, fill=X, expand=True, anchor='n')
                fr7 = Frame(fr0, bg='#1d366c')
                fr7.pack(side=BOTTOM, fill=X, expand=True, anchor='n', pady=8)

                image_entrada = Image.open('imagens\\cadastro.png')
                resize_entrada = image_entrada.resize((30, 30))
                nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                lbl = Label(fr1, text=" Cadastro de Produtos", image=nova_image_entrada, font=fonte_titulos, compound="left", bg="#ffffff", fg='#000000')
                lbl.grid(row=0, column=1)
                fr1.grid_columnconfigure(0, weight=1)
                fr1.grid_columnconfigure(2, weight=1)

                # fr2 linha horizontal

                lblnome_produto = Label(fr3, text="Nome do Produto:", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblnome_produto.grid(row=1, column=1, pady=4)
                entrynome_produto = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entrynome_produto.grid(row=1, column=2, pady=4)

                lblest_minimo = Label(fr3, text="Estoque Mínimo:", font=fonte_padrao, fg='#000000', bg='#ffffff')
                lblest_minimo.grid(row=2, column=1, pady=4)
                entryest_minimo = Entry(fr3, font=fonte_padrao, fg='#000000', justify='center', width=40)
                entryest_minimo.grid(row=2, column=2, pady=4)

                btn_imagem = Button(fr3, text='Adicionar Imagem', highlightthickness=0, width=20, relief=RIDGE,
                                      command=upload_imagem, font=fonte_botao, cursor="hand2")
                btn_imagem.grid(row=3, column=1, columnspan=2, pady=8)

                fr3.grid_columnconfigure(0, weight=1)
                fr3.grid_columnconfigure(3, weight=1)

                # fr4 linha

                btn_cadastro = Button(fr5, text='Cadastrar', bg='#1d366c', fg='#FFFFFF',
                                      activebackground='#1d366c',
                                      activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                                      command=salvar, font=fonte_botao, cursor="hand2")
                btn_cadastro.grid(row=0, column=1, pady=4)

                fr5.grid_columnconfigure(0, weight=1)
                fr5.grid_columnconfigure(2, weight=1)

                style = ttk.Style()
                # style.theme_use('default')
                style.configure('Treeview',
                                background='#ffffff',
                                rowheight=24,
                                fieldbackground='#ffffff',
                                font=fonte_padrao)
                style.configure("Treeview.Heading",
                                foreground='#000000',
                                background="#ffffff",
                                font=fonte_padrao)
                style.map('Treeview', background=[('selected', '#1d366c')])

                tree_principal = ttk.Treeview(frame4, selectmode='none')
                vsb = ttk.Scrollbar(frame4, orient="vertical", command=tree_principal.yview)
                vsb.pack(side=RIGHT, fill='y')
                tree_principal.configure(yscrollcommand=vsb.set)
                vsbx = ttk.Scrollbar(frame4, orient="horizontal", command=tree_principal.xview)
                vsbx.pack(side=BOTTOM, fill='x')
                tree_principal.configure(xscrollcommand=vsbx.set)
                tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
                tree_principal["columns"] = ("1", "2", "3", "4")
                tree_principal['show'] = 'headings'
                tree_principal.column("1", width=100, anchor='c')
                tree_principal.column("2", width=100, anchor='c')
                tree_principal.column("3", width=100, anchor='c')
                tree_principal.column("4", width=100, anchor='c')
                tree_principal.heading("1", text="Nome do Produto")
                tree_principal.heading("2", text="Nome da Imagem")
                tree_principal.heading("3", text="Quantidade em Estoque")
                tree_principal.heading("4", text="Estoque Mínimo")
                tree_principal.tag_configure('par', background='#e9e9e9')
                tree_principal.tag_configure('impar', background='#ffffff')
                atualizar_lista()
                root3.mainloop()

            frame0 = Frame(root3, bg='#ffffff')
            frame0.grid(row=0, column=0, stick='nsew')
            root3.grid_rowconfigure(0, weight=1)
            root3.grid_columnconfigure(0, weight=1)
            frame1 = Frame(frame0, bg="#1d366c")
            frame1.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame2 = Frame(frame0, bg='#ffffff', pady=6)
            frame2.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame3 = Frame(frame0, bg='#1d366c')
            frame3.pack(side=TOP, fill=X, expand=False, anchor='n')
            frame4 = Frame(frame0, bg='#ffffff', pady=10)
            frame4.pack(side=TOP, fill=BOTH, expand=True, anchor='n')
            frame5 = Frame(frame0, bg='#1d366c')
            frame5.pack(side=TOP, fill=X, expand=False, anchor='n')


            image_estoque_titulo = Image.open('imagens\\estoque_titulo.png')
            resize_estoque_titulo = image_estoque_titulo.resize((35, 35))
            nova_image_estoque_titulo = ImageTk.PhotoImage(resize_estoque_titulo)

            lbllogin = Label(frame1, image=nova_image_estoque_titulo, text=" Estoque TI", compound="left", bg='#1d366c',
                             fg='#FFFFFF', font=fonte_titulos)
            lbllogin.photo = nova_image_estoque_titulo
            lbllogin.grid(row=0, column=1)
            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(2, weight=1)

            def muda_home(e):
                image_home = Image.open('imagens\\home_over.png')
                resize_home = image_home.resize((30, 30))
                nova_image_home = ImageTk.PhotoImage(resize_home)
                btnhome.photo = nova_image_home
                btnhome.config(image=nova_image_home, fg='#7c7c7c')

            def volta_home(e):
                image_home = Image.open('imagens\\home.png')
                resize_home = image_home.resize((30, 30))
                nova_image_home = ImageTk.PhotoImage(resize_home)
                btnhome.photo = nova_image_home
                btnhome.config(image=nova_image_home, fg='#000000')

            image_home = Image.open('imagens\\home.png')
            resize_home = image_home.resize((30, 30))
            nova_image_home = ImageTk.PhotoImage(resize_home)
            btnhome = Button(frame2, image=nova_image_home, text=" Home", compound="left",
                                   font=fonte_titulos, bg="#ffffff", fg='#000000', command=home, borderwidth=0,
                                   relief=RIDGE,
                                   activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btnhome.photo = nova_image_home
            btnhome.grid(row=0, column=1, padx=15)
            btnhome.bind("<Enter>", muda_home)
            btnhome.bind("<Leave>", volta_home)


            def muda_entrada(e):
                image_entrada = Image.open('imagens\\entrada_over.png')
                resize_entrada = image_entrada.resize((25, 30))
                nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                btnentrada.photo = nova_image_entrada
                btnentrada.config(image=nova_image_entrada, fg='#7c7c7c')

            def volta_entrada(e):
                image_entrada = Image.open('imagens\\entrada.png')
                resize_entrada = image_entrada.resize((25, 30))
                nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
                btnentrada.photo = nova_image_entrada
                btnentrada.config(image=nova_image_entrada, fg='#000000')

            image_entrada = Image.open('imagens\\entrada.png')
            resize_entrada = image_entrada.resize((25, 30))
            nova_image_entrada = ImageTk.PhotoImage(resize_entrada)
            btnentrada = Button(frame2, image=nova_image_entrada, text=" Entrada", compound="left",
                                   font=fonte_titulos, bg="#ffffff", fg='#000000', command=entrada, borderwidth=0,
                                   relief=RIDGE,
                                   activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btnentrada.photo = nova_image_entrada
            btnentrada.grid(row=0, column=2, padx=15)
            btnentrada.bind("<Enter>", muda_entrada)
            btnentrada.bind("<Leave>", volta_entrada)

            def muda_saida(e):
                image_saida = Image.open('imagens\\saida_over.png')
                resize_saida = image_saida.resize((25, 30))
                nova_image_saida = ImageTk.PhotoImage(resize_saida)
                btnsaida.photo = nova_image_saida
                btnsaida.config(image=nova_image_saida, fg='#7c7c7c')

            def volta_saida(e):
                image_saida = Image.open('imagens\\saida.png')
                resize_saida = image_saida.resize((25, 30))
                nova_image_saida = ImageTk.PhotoImage(resize_saida)
                btnsaida.photo = nova_image_saida
                btnsaida.config(image=nova_image_saida, fg='#000000')

            image_saida = Image.open('imagens\\saida.png')
            resize_saida = image_saida.resize((25, 30))
            nova_image_saida = ImageTk.PhotoImage(resize_saida)
            btnsaida = Button(frame2, image=nova_image_saida, text=" Saída", compound="left",
                                   font=fonte_titulos, bg="#ffffff", fg='#000000', command=saida, borderwidth=0,
                                   relief=RIDGE,
                                   activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btnsaida.photo = nova_image_saida
            btnsaida.grid(row=0, column=3, padx=15)
            btnsaida.bind("<Enter>", muda_saida)
            btnsaida.bind("<Leave>", volta_saida)


            def muda_historico(e):
                image_historico = Image.open('imagens\\historico_over.png')
                resize_historico = image_historico.resize((25, 30))
                nova_image_historico = ImageTk.PhotoImage(resize_historico)
                btnhistorico.photo = nova_image_historico
                btnhistorico.config(image=nova_image_historico, fg='#7c7c7c')

            def volta_historico(e):
                image_historico = Image.open('imagens\\historico.png')
                resize_historico = image_historico.resize((25, 30))
                nova_image_historico = ImageTk.PhotoImage(resize_historico)
                btnhistorico.photo = nova_image_historico
                btnhistorico.config(image=nova_image_historico, fg='#000000')

            image_historico = Image.open('imagens\\historico.png')
            resize_historico = image_historico.resize((25, 30))
            nova_image_historico = ImageTk.PhotoImage(resize_historico)
            btnhistorico = Button(frame2, image=nova_image_historico, text=" Histórico", compound="left",
                                   font=fonte_titulos, bg="#ffffff", fg='#000000', command=historico, borderwidth=0,
                                   relief=RIDGE,
                                   activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btnhistorico.photo = nova_image_historico
            btnhistorico.grid(row=0, column=4, padx=15)
            btnhistorico.bind("<Enter>", muda_historico)
            btnhistorico.bind("<Leave>", volta_historico)

            def muda_cadastro(e):
                image_cadastro = Image.open('imagens\\cadastro_over.png')
                resize_cadastro = image_cadastro.resize((30, 30))
                nova_image_cadastro = ImageTk.PhotoImage(resize_cadastro)
                btncadastro.photo = nova_image_cadastro
                btncadastro.config(image=nova_image_cadastro, fg='#7c7c7c')

            def volta_cadastro(e):
                image_cadastro = Image.open('imagens\\cadastro.png')
                resize_cadastro = image_cadastro.resize((30, 30))
                nova_image_cadastro = ImageTk.PhotoImage(resize_cadastro)
                btncadastro.photo = nova_image_cadastro
                btncadastro.config(image=nova_image_cadastro, fg='#000000')

            image_cadastro = Image.open('imagens\\cadastro.png')
            resize_cadastro = image_cadastro.resize((30, 30))
            nova_image_cadastro = ImageTk.PhotoImage(resize_cadastro)
            btncadastro = Button(frame2, image=nova_image_cadastro, text=" Cadastro", compound="left",
                                   font=fonte_titulos, bg="#ffffff", fg='#000000', command=cadastro, borderwidth=0,
                                   relief=RIDGE,
                                   activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btncadastro.photo = nova_image_cadastro
            btncadastro.grid(row=0, column=5, padx=15)
            btncadastro.bind("<Enter>", muda_cadastro)
            btncadastro.bind("<Leave>", volta_cadastro)


            def muda_controle(e):
                image_controle = Image.open('imagens\\controle_over.png')
                resize_controle = image_controle.resize((25, 30))
                nova_image_controle = ImageTk.PhotoImage(resize_controle)
                btncontrole.photo = nova_image_controle
                btncontrole.config(image=nova_image_controle, fg='#7c7c7c')

            def volta_controle(e):
                image_controle = Image.open('imagens\\controle.png')
                resize_controle = image_controle.resize((25, 30))
                nova_image_controle = ImageTk.PhotoImage(resize_controle)
                btncontrole.photo = nova_image_controle
                btncontrole.config(image=nova_image_controle, fg='#000000')

            image_controle = Image.open('imagens\\controle.png')
            resize_controle = image_controle.resize((25, 30))
            nova_image_controle = ImageTk.PhotoImage(resize_controle)
            btncontrole = Button(frame2, image=nova_image_controle, text=" Controle", compound="left",
                                   font=fonte_titulos, bg="#ffffff", fg='#000000', command="controle", borderwidth=0,
                                   relief=RIDGE,
                                   activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
            btncontrole.photo = nova_image_controle
            btncontrole.grid(row=0, column=6, padx=15)
            btncontrole.bind("<Enter>", muda_controle)
            btncontrole.bind("<Leave>", volta_controle)

            frame2.grid_columnconfigure(0, weight=1)
            frame2.grid_columnconfigure(7, weight=1)

            # frame3 linha horizontal.

            # frame4 Frame com conteudo alternado.

            Label(frame5, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1, sticky="ew")
            frame5.grid_columnconfigure(0, weight=1)
            frame5.grid_columnconfigure(2, weight=1)

            root3.state('zoomed')
            home()
        def config():
            root3 = Toplevel()
            root3.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
            root3.unbind_class("Button", "<Key-space>")
            root3.focus_force()
            root3.grab_set()
            root3.state('zoomed')

            def home():
                def permissao():
                    if usuariologado != 'Administrador':
                        btncad_user.config(state='disabled')

                def cadastro_usuario():
                    def salvar():
                        nome = ent_nome.get()
                        login = ent_login.get()
                        senha = ent_senha.get()
                        confirma_senha = ent_confir_senha.get()
                        if nome == "" or senha == "" or confirma_senha == "":
                            messagebox.showwarning('Cadastro de Usuários:', 'Todos os campos devem ser preenchidos.',
                                                   parent=root3)
                        elif senha != confirma_senha:
                            messagebox.showwarning('Cadastro de Usuários:', 'As senhas não conferem!', parent=root3)
                        else:
                            r = cursor.execute("SELECT * FROM dbo.analista WHERE login=?", (login,))
                            result = r.fetchone()
                            if result == None:
                                hashed = bcrypt.hashpw(confirma_senha.encode("utf-8"), bcrypt.gensalt())
                                try:
                                    cursor.execute(
                                        "INSERT INTO dbo.analista (nome_analista, login, senha) values(?,?,?)",
                                        (nome, login, hashed.decode("utf-8")))
                                    cursor.commit()
                                except:
                                    messagebox.showerror('Banco de Dados:', 'Erro de conexão com o Banco de Dados.',
                                                         parent=root3)
                                    return False
                                messagebox.showwarning('Cadastro de Usuários:', 'Usuário cadastrado com sucesso!',
                                                       parent=root3)
                                ent_nome.delete(0, END)
                                ent_login.delete(0, END)
                                ent_senha.delete(0, END)
                                ent_confir_senha.delete(0, END)
                                ent_nome.focus_force()
                            else:
                                messagebox.showwarning('Cadastro de Usuários:', 'Login já cadastrado!', parent=root3)

                    def salvar_bind(event):
                        salvar()

                    for widget in frame3.winfo_children():
                        widget.destroy()

                    fr1 = Frame(frame3, bg='#ffffff')
                    fr1.pack(side=TOP, fill=X)
                    fr2 = Frame(frame3, bg='#1d366c')  # linha
                    fr2.pack(side=TOP, fill=X, pady=4)
                    fr3 = Frame(frame3, bg='#ffffff')
                    fr3.pack(side=TOP, fill=X)
                    fr4 = Frame(frame3, bg='#ffffff')
                    fr4.pack(side=TOP, fill=X)


                    lbl = Label(fr1, text=" Cadastro de Usuários:", font=fonte_titulos, fg='#000000', bg='#ffffff')
                    lbl.grid(row=0, column=1)
                    fr1.grid_columnconfigure(0, weight=1)
                    fr1.grid_columnconfigure(2, weight=1)

                    # fr2 linha horizontal

                    lbl_nome = Label(fr3, text="Nome Completo:", font=fonte_padrao, fg='#222222', bg='#ffffff')
                    lbl_nome.grid(row=1, column=1, pady=8)
                    ent_nome = Entry(fr3, font=fonte_padrao, fg='#222222', width=30, highlightthickness=1,
                                     highlightbackground="#5f5d5d", highlightcolor="#5f5d5d", relief=FLAT)
                    ent_nome.bind("<Return>", salvar_bind)
                    ent_nome.grid(row=1, column=2)
                    ent_nome.focus_force()

                    lbl_login = Label(fr3, text="Login:", font=fonte_padrao, fg='#222222', bg='#ffffff')
                    lbl_login.grid(row=2, column=1, pady=8)
                    ent_login = Entry(fr3, font=fonte_padrao, fg='#222222', width=30, highlightthickness=1,
                                      highlightbackground="#5f5d5d", highlightcolor="#5f5d5d", relief=FLAT)
                    ent_login.bind("<Return>", salvar_bind)
                    ent_login.grid(row=2, column=2)

                    lbl_senha = Label(fr3, text="Senha:", font=fonte_padrao, fg='#222222', bg='#ffffff')
                    lbl_senha.grid(row=3, column=1, pady=8)
                    ent_senha = Entry(fr3, font=fonte_padrao, fg='#222222', width=30, show='*', highlightthickness=1,
                                      highlightbackground="#5f5d5d", highlightcolor="#5f5d5d", relief=FLAT)
                    ent_senha.bind("<Return>", salvar_bind)
                    ent_senha.grid(row=3, column=2)

                    lbl_confir_senha = Label(fr3, text="Confirmar Senha:", font=fonte_padrao, fg='#222222',
                                             bg='#ffffff')
                    lbl_confir_senha.grid(row=4, column=1, pady=8)
                    ent_confir_senha = Entry(fr3, font=fonte_padrao, fg='#222222', width=30, show='*',
                                             highlightthickness=1,
                                             highlightbackground="#5f5d5d", highlightcolor="#5f5d5d", relief=FLAT)
                    ent_confir_senha.bind("<Return>", salvar_bind)
                    ent_confir_senha.grid(row=4, column=2)

                    fr3.grid_columnconfigure(0, weight=1)
                    fr3.grid_columnconfigure(3, weight=1)

                    btn_cadastro = Button(fr4, text='Salvar', bg='#1d366c', fg='#FFFFFF',
                                          activebackground='#1d366c',
                                          activeforeground="#FFFFFF", highlightthickness=0, width=20, relief=RIDGE,
                                          command=salvar,
                                          font=fonte_botao, cursor="hand2")
                    btn_cadastro.grid(row=0, column=1, padx=6)

                    fr4.grid_columnconfigure(0, weight=1)
                    fr4.grid_columnconfigure(2, weight=1)


                for widget in frame3.winfo_children():
                    widget.destroy()
                fr0 = Frame(frame3, bg='#ffffff')
                fr0.pack(side=TOP, fill=X)

                def muda_cad_user(e):
                    image_cad_user = Image.open('imagens\\cadastro_usuario.png')
                    resize_cad_user = image_cad_user.resize((40, 40))
                    nova_image_cad_user = ImageTk.PhotoImage(resize_cad_user)
                    btncad_user.photo = nova_image_cad_user
                    btncad_user.config(image=nova_image_cad_user, fg='#1d366c')
                def volta_cad_user(e):
                    image_cad_user = Image.open('imagens\\cadastro_usuario.png')
                    resize_cad_user = image_cad_user.resize((40, 40))
                    nova_image_cad_user = ImageTk.PhotoImage(resize_cad_user)
                    btncad_user.photo = nova_image_cad_user
                    btncad_user.config(image=nova_image_cad_user, fg='#000000')

                image_cad_user = Image.open('imagens\\cadastro_usuario.png')
                resize_cad_user = image_cad_user.resize((40, 40))
                nova_image_cad_user = ImageTk.PhotoImage(resize_cad_user)
                btncad_user = Button(fr0, image=nova_image_cad_user, text=" Cadastro de Usuários", compound="top",
                                    font=fonte_titulos, bg="#ffffff", fg='#000000', command=cadastro_usuario, borderwidth=0,
                                    relief=RIDGE,
                                    activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
                btncad_user.photo = nova_image_cad_user
                btncad_user.grid(row=0, column=1, padx=5)
                btncad_user.bind("<Enter>", muda_cad_user)
                btncad_user.bind("<Leave>", volta_cad_user)


                def muda_cad_setor(e):
                    image_cad_setor = Image.open('imagens\\cadastro_usuario.png')
                    resize_cad_setor = image_cad_setor.resize((40, 40))
                    nova_image_cad_setor = ImageTk.PhotoImage(resize_cad_setor)
                    btncad_setor.photo = nova_image_cad_setor
                    btncad_setor.config(image=nova_image_cad_setor, fg='#1d366c')
                def volta_cad_setor(e):
                    image_cad_setor = Image.open('imagens\\cadastro_usuario.png')
                    resize_cad_setor = image_cad_setor.resize((40, 40))
                    nova_image_cad_setor = ImageTk.PhotoImage(resize_cad_setor)
                    btncad_setor.photo = nova_image_cad_setor
                    btncad_setor.config(image=nova_image_cad_setor, fg='#000000')

                image_cad_setor = Image.open('imagens\\cadastro_usuario.png')
                resize_cad_setor = image_cad_setor.resize((40, 40))
                nova_image_cad_setor = ImageTk.PhotoImage(resize_cad_setor)
                btncad_setor = Button(fr0, image=nova_image_cad_setor, text=" Cadastro de Setores", compound="top",
                                    font=fonte_titulos, bg="#ffffff", fg='#000000', command="cad_setor", borderwidth=0,
                                    relief=RIDGE,
                                    activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
                btncad_setor.photo = nova_image_cad_setor
                btncad_setor.grid(row=0, column=2, padx=5)
                btncad_setor.bind("<Enter>", muda_cad_setor)
                btncad_setor.bind("<Leave>", volta_cad_setor)

                fr0.grid_columnconfigure(0, weight=1)
                fr0.grid_columnconfigure(3, weight=1)
                permissao()
                root3.mainloop()

            frame1 = Frame(root3, bg="#1d366c")
            frame1.pack(side=TOP, fill=X, expand=False)
            frame2 = Frame(root3, bg='#1d366c')
            frame2.pack(side=TOP, fill=X, expand=False, pady=4)
            frame3 = Frame(root3, bg='#ffffff')
            frame3.pack(side=TOP, fill=BOTH, expand=True)
            frame4 = Frame(root3, bg='#1d366c', height=20)
            frame4.pack(side=TOP, fill=X, expand=False)

            image_estoque_titulo = Image.open('imagens\\config.png')
            resize_estoque_titulo = image_estoque_titulo.resize((55, 55))
            nova_image_estoque_titulo = ImageTk.PhotoImage(resize_estoque_titulo)

            lbllogin = Label(frame1, image=nova_image_estoque_titulo, text=" Configurações", compound="left", bg='#1d366c',
                             fg='#FFFFFF', font=fonte_titulos)
            lbllogin.photo = nova_image_estoque_titulo
            lbllogin.grid(row=0, column=1)
            frame1.grid_columnconfigure(0, weight=1)
            frame1.grid_columnconfigure(2, weight=1)
            home()


        global lbl_loop
        lbl_loop = Label(frame4, text='', bg="#1d366c")
        lbl_loop.grid(row=0, column=2)

        root2 = Toplevel()
        root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
        root2.unbind_class("Button", "<Key-space>")
        root2.focus_force()
        root2.grab_set()

        frame0 = Frame(root2, bg='#ffffff')
        frame0.grid(row=0, column=0, stick='nsew')
        root2.grid_rowconfigure(0, weight=1)
        root2.grid_columnconfigure(0, weight=1)
        frame1 = Frame(frame0, bg="#1d366c")
        frame1.pack(side=TOP, fill=X, expand=False, anchor='center', pady=(0,20))
        frame2 = Frame(frame0, bg='#ffffff')
        frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=6)
        frame3 = Frame(frame0, bg='#1d366c')
        frame3.pack(side=TOP, fill=X, expand=False, anchor='center', pady=(20,0))

        lblferr = Label(frame1, image=nova_image_ferramentas, text=" Ferramentas", compound="left", bg='#1d366c',
                         fg='#FFFFFF', font=fonte_titulos)
        lblferr.grid(row=0, column=1)
        frame1.grid_columnconfigure(0, weight=1)
        frame1.grid_columnconfigure(2, weight=1)

        def muda_trocasenha(e):
            image_trocasenha = Image.open('imagens\\trocasenha_over.png')
            resize_trocasenha = image_trocasenha.resize((55, 55))
            nova_image_trocasenha = ImageTk.PhotoImage(resize_trocasenha)
            btntrocasenha.photo = nova_image_trocasenha
            btntrocasenha.config(image=nova_image_trocasenha, fg='#7c7c7c')
        def volta_trocasenha(e):
            image_trocasenha = Image.open('imagens\\trocasenha.png')
            resize_trocasenha = image_trocasenha.resize((55, 55))
            nova_image_trocasenha = ImageTk.PhotoImage(resize_trocasenha)
            btntrocasenha.photo = nova_image_trocasenha
            btntrocasenha.config(image=nova_image_trocasenha, fg='#1d366c')
        image_trocasenha = Image.open('imagens\\trocasenha.png')
        resize_trocasenha = image_trocasenha.resize((55, 55))
        nova_image_trocasenha = ImageTk.PhotoImage(resize_trocasenha)
        btntrocasenha = Button(frame2, image=nova_image_trocasenha, text="Alterar Senha", compound="top",
                          font=fonte_titulos, bg="#fff3ff", fg='#1d366c', command=trocasenha, borderwidth=0,
                          relief=RIDGE,
                          activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
        btntrocasenha.photo = nova_image_trocasenha
        btntrocasenha.grid(row=0, column=1, padx=15)
        btntrocasenha.bind("<Enter>", muda_trocasenha)
        btntrocasenha.bind("<Leave>", volta_trocasenha)

        def muda_relatorio(e):
            image_relatorio = Image.open('imagens\\relatorio_over.png')
            resize_relatorio = image_relatorio.resize((55, 55))
            nova_image_relatorio = ImageTk.PhotoImage(resize_relatorio)
            btnrelatorio.photo = nova_image_relatorio
            btnrelatorio.config(image=nova_image_relatorio, fg='#7c7c7c')
        def volta_relatorio(e):
            image_relatorio = Image.open('imagens\\relatorio.png')
            resize_relatorio = image_relatorio.resize((55, 55))
            nova_image_relatorio = ImageTk.PhotoImage(resize_relatorio)
            btnrelatorio.photo = nova_image_relatorio
            btnrelatorio.config(image=nova_image_relatorio, fg='#1d366c')
        image_relatorio = Image.open('imagens\\relatorio.png')
        resize_relatorio = image_relatorio.resize((55, 55))
        nova_image_relatorio = ImageTk.PhotoImage(resize_relatorio)
        btnrelatorio = Button(frame2, image=nova_image_relatorio, text="Relatório", compound="top",
                          font=fonte_titulos, bg="#ffffff", fg='#1d366c', command=relatorio, borderwidth=0,
                          relief=RIDGE,
                          activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
        btnrelatorio.photo = nova_image_relatorio
        btnrelatorio.grid(row=0, column=2, padx=15)
        btnrelatorio.bind("<Enter>", muda_relatorio)
        btnrelatorio.bind("<Leave>", volta_relatorio)


        def muda_estoque(e):
            image_estoque = Image.open('imagens\\estoque_over.png')
            resize_estoque = image_estoque.resize((55, 55))
            nova_image_estoque = ImageTk.PhotoImage(resize_estoque)
            btnestoque.photo = nova_image_estoque
            btnestoque.config(image=nova_image_estoque, fg='#7c7c7c')
        def volta_estoque(e):
            image_estoque = Image.open('imagens\\estoque.png')
            resize_estoque = image_estoque.resize((55, 55))
            nova_image_estoque = ImageTk.PhotoImage(resize_estoque)
            btnestoque.photo = nova_image_estoque
            btnestoque.config(image=nova_image_estoque, fg='#1d366c')
        image_estoque = Image.open('imagens\\estoque.png')
        resize_estoque = image_estoque.resize((55, 55))
        nova_image_estoque = ImageTk.PhotoImage(resize_estoque)
        btnestoque = Button(frame2, image=nova_image_estoque, text="Estoque TI", compound="top",
                          font=fonte_titulos, bg="#fff3ff", fg='#1d366c', command=estoque, borderwidth=0,
                          relief=RIDGE,
                          activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
        btnestoque.photo = nova_image_estoque
        btnestoque.grid(row=0, column=3, padx=15)
        btnestoque.bind("<Enter>", muda_estoque)
        btnestoque.bind("<Leave>", volta_estoque)


        def muda_config(e):
            image_config = Image.open('imagens\\config_over.png')
            resize_config = image_config.resize((55, 55))
            nova_image_config = ImageTk.PhotoImage(resize_config)
            btnconfig.photo = nova_image_config
            btnconfig.config(image=nova_image_config, fg='#7c7c7c')
        def volta_config(e):
            image_config = Image.open('imagens\\config.png')
            resize_config = image_config.resize((55, 55))
            nova_image_config = ImageTk.PhotoImage(resize_config)
            btnconfig.photo = nova_image_config
            btnconfig.config(image=nova_image_config, fg='#1d366c')
        image_config = Image.open('imagens\\config.png')
        resize_config = image_config.resize((55, 55))
        nova_image_config = ImageTk.PhotoImage(resize_config)
        btnconfig = Button(frame2, image=nova_image_config, text="Configurações", compound="top",
                          font=fonte_titulos, bg="#fff3ff", fg='#1d366c', command=config, borderwidth=0,
                          relief=RIDGE,
                          activebackground="#ffffff", activeforeground="#7c7c7c", cursor="hand2")
        btnconfig.photo = nova_image_config
        btnconfig.grid(row=0, column=4)
        btnconfig.bind("<Enter>", muda_config)
        btnconfig.bind("<Leave>", volta_config)

        frame2.grid_columnconfigure(0, weight=1)
        frame2.grid_columnconfigure(5, weight=1)

        Label(frame3, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1, sticky="ew")
        frame3.grid_columnconfigure(0, weight=1)
        frame3.grid_columnconfigure(2, weight=1)
        '''root2.update()
        largura = frame0.winfo_width()
        altura = frame0.winfo_height()
        print(largura, altura)'''
        window_width = 486
        window_height = 198
        screen_width = root2.winfo_screenwidth()
        screen_height = root2.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
        #root2.resizable(0, 0)
        root2.iconbitmap('imagens\\ico.ico')
        root2.title(titulo_todos)
        permissao()
        root2.mainloop()
    # /////////////////////////////FIM FERRAMENTAS/////////////////////////////
    def loop_atualização():
        if controle_loop == 0:
            lbl_loop.after(300000, atualizar_lista_principal)

    # -------------------FRAME PRINCIPAL-------------------#
    frame0 = Frame(root, bg="#1d366c")
    frame0.grid(row=0, column=0, stick='nsew')
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    frame1 = Frame(frame0, bg="#1d366c")
    frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
    frame2 = Frame(frame0, bg="#1d366c")
    frame2.pack(side=TOP, fill=X, expand=False, anchor='n')
    frame3 = Frame(frame0, highlightbackground="#2c2c2c", highlightcolor="#2c2c2c", highlightthickness=1, borderwidth=2)
    frame3.pack(side=TOP, fill=BOTH, expand=True, anchor='n')
    frame4 = Frame(frame0, bg="#1d366c")
    frame4.pack(side=TOP, fill=X, expand=False, anchor='n')

    # -------------------FRAME1-------------------#
    image_logo = Image.open('imagens\\logo.png')
    resize_logo = image_logo.resize((469, 106))
    nova_image_logo = ImageTk.PhotoImage(resize_logo)
    lbl1 = Label(frame1, image=nova_image_logo, bg="#1d366c")
    lbl1.photo = nova_image_logo
    lbl1.grid(row=0, column=1)
    frame1.grid_columnconfigure(0, weight=1)
    frame1.grid_columnconfigure(2, weight=1)

    # -------------------FRAME2-------------------#
    def muda_login(e):
        image_login = Image.open('imagens\\login_over.png')
        resize_login = image_login.resize((35, 35))
        nova_image_login = ImageTk.PhotoImage(resize_login)
        btnlogin.photo = nova_image_login
        btnlogin.config(image=nova_image_login, fg='#7c7c7c')
    def volta_login(e):
        image_login = Image.open('imagens\\login.png')
        resize_login = image_login.resize((35, 35))
        nova_image_login = ImageTk.PhotoImage(resize_login)
        btnlogin.photo = nova_image_login
        btnlogin.config(image=nova_image_login, fg='#ffffff')
    image_login = Image.open('imagens\\login.png')
    resize_login = image_login.resize((35, 35))
    nova_image_login = ImageTk.PhotoImage(resize_login)
    btnlogin = Button(frame2, image=nova_image_login, text=" Trocar Usuário", compound="left",
                      font=fonte_padrao, bg="#1d366c", fg='#FFFFFF', command=login_interno, borderwidth=0, relief=RIDGE,
                      activebackground="#1d366c", activeforeground="#1d366c", cursor="hand2")
    btnlogin.photo = nova_image_login
    btnlogin.grid(row=0, column=1, pady=6, padx=10)
    btnlogin.bind("<Enter>", muda_login)
    btnlogin.bind("<Leave>", volta_login)

    def muda_chamado(e):
        image_chamado = Image.open('imagens\\chamado_over.png')
        resize_chamado = image_chamado.resize((30, 35))
        nova_image_chamado = ImageTk.PhotoImage(resize_chamado)
        btnchamado.photo = nova_image_chamado
        btnchamado.config(image=nova_image_chamado, fg='#7c7c7c')
    def volta_chamado(e):
        image_chamado = Image.open('imagens\\chamado.png')
        resize_chamado = image_chamado.resize((30, 35))
        nova_image_chamado = ImageTk.PhotoImage(resize_chamado)
        btnchamado.photo = nova_image_chamado
        btnchamado.config(image=nova_image_chamado, fg='#ffffff')

    image_chamado = Image.open('imagens\\chamado.png')
    resize_chamado = image_chamado.resize((30, 35))
    nova_image_chamado = ImageTk.PhotoImage(resize_chamado)
    btnchamado = Button(frame2, image=nova_image_chamado, text=" + Abrir Chamado (F1)", compound="left", font=fonte_padrao,
                        bg="#1d366c", fg='#FFFFFF', command=abrirchamado, borderwidth=0, relief=RIDGE,
                        activebackground="#1d366c", activeforeground="#1d366c", cursor="hand2")
    btnchamado.photo = nova_image_chamado
    btnchamado.grid(row=0, column=2, pady=6, padx=10)
    btnchamado.bind("<Enter>", muda_chamado)
    btnchamado.bind("<Leave>", volta_chamado)

    def muda_atendimento(e):
        image_atendimento = Image.open('imagens\\atendimento_over.png')
        resize_atendimento = image_atendimento.resize((30, 35))
        nova_image_atendimento = ImageTk.PhotoImage(resize_atendimento)
        btnatendimento.photo = nova_image_atendimento
        btnatendimento.config(image=nova_image_atendimento, fg='#7c7c7c')
    def volta_atendimento(e):
        image_atendimento = Image.open('imagens\\atendimento.png')
        resize_atendimento = image_atendimento.resize((30, 35))
        nova_image_atendimento = ImageTk.PhotoImage(resize_atendimento)
        btnatendimento.photo = nova_image_atendimento
        btnatendimento.config(image=nova_image_atendimento, fg='#ffffff')

    image_atendimento = Image.open('imagens\\atendimento.png')
    resize_atendimento = image_atendimento.resize((30, 35))
    nova_image_atendimento = ImageTk.PhotoImage(resize_atendimento)
    btnatendimento = Button(frame2, image=nova_image_atendimento, text=" Atendimento (F2)", compound="left",
                            font=fonte_padrao, bg="#1d366c", fg='#FFFFFF', command=atendimento, borderwidth=0,
                            relief=RIDGE, activebackground="#1d366c", activeforeground="#1d366c", cursor="hand2")
    btnatendimento.photo = nova_image_atendimento
    btnatendimento.grid(row=0, column=3, pady=6, padx=10)
    btnatendimento.bind("<Enter>", muda_atendimento)
    btnatendimento.bind("<Leave>", volta_atendimento)

    def muda_visualizarchamado(e):
        image_visualizarchamado = Image.open('imagens\\visualizar_over.png')
        resize_visualizarchamado = image_visualizarchamado.resize((30, 25))
        nova_image_visualizarchamado = ImageTk.PhotoImage(resize_visualizarchamado)
        btnvisualizarchamado.photo = nova_image_visualizarchamado
        btnvisualizarchamado.config(image=nova_image_visualizarchamado, fg='#7c7c7c')
    def volta_visualizarchamado(e):
        image_visualizarchamado = Image.open('imagens\\visualizar.png')
        resize_visualizarchamado = image_visualizarchamado.resize((30, 25))
        nova_image_visualizarchamado = ImageTk.PhotoImage(resize_visualizarchamado)
        btnvisualizarchamado.photo = nova_image_visualizarchamado
        btnvisualizarchamado.config(image=nova_image_visualizarchamado, fg='#ffffff')

    image_visualizarchamado = Image.open('imagens\\visualizar.png')
    resize_visualizarchamado = image_visualizarchamado.resize((30, 25))
    nova_image_visualizarchamado = ImageTk.PhotoImage(resize_visualizarchamado)
    btnvisualizarchamado = Button(frame2, image=nova_image_visualizarchamado, text=" Visualizar\Editar Chamado (F3)",
                                  compound="left", font=fonte_padrao, bg="#1d366c", fg='#FFFFFF', command=visualizar_chamado,
                                  borderwidth=0, relief=RIDGE, activebackground="#1d366c", activeforeground="#1d366c", cursor="hand2")
    btnvisualizarchamado.photo = nova_image_visualizarchamado
    btnvisualizarchamado.grid(row=0, column=4, pady=6, padx=10)
    btnvisualizarchamado.bind("<Enter>", muda_visualizarchamado)
    btnvisualizarchamado.bind("<Leave>", volta_visualizarchamado)

    def muda_ferramentas(e):
        image_ferramentas = Image.open('imagens\\ferramentas_over.png')
        resize_ferramentas = image_ferramentas.resize((35, 35))
        nova_image_ferramentas = ImageTk.PhotoImage(resize_ferramentas)
        btnferramentas.photo = nova_image_ferramentas
        btnferramentas.config(image=nova_image_ferramentas, fg='#7c7c7c')
    def volta_ferramentas(e):
        image_ferramentas = Image.open('imagens\\ferramentas.png')
        resize_ferramentas = image_ferramentas.resize((35, 35))
        nova_image_ferramentas = ImageTk.PhotoImage(resize_ferramentas)
        btnferramentas.photo = nova_image_ferramentas
        btnferramentas.config(image=nova_image_ferramentas, fg='#ffffff')

    image_ferramentas = Image.open('imagens\\ferramentas.png')
    resize_ferramentas = image_ferramentas.resize((35, 35))
    nova_image_ferramentas = ImageTk.PhotoImage(resize_ferramentas)
    btnferramentas = Button(frame2, image=nova_image_ferramentas, text=" Ferramentas (F4)", compound="left",
                            font=fonte_padrao, bg="#1d366c", fg='#FFFFFF', command=ferramentas,
                            borderwidth=0, relief=RIDGE, activebackground="#1d366c", activeforeground="#7c7c7c", cursor="hand2")
    btnferramentas.photo = nova_image_ferramentas
    btnferramentas.grid(row=0, column=5, pady=6, padx=10)
    btnferramentas.bind("<Enter>", muda_ferramentas)
    btnferramentas.bind("<Leave>", volta_ferramentas)

    clique_busca = StringVar()
    drop_busca = OptionMenu(frame2, clique_busca, 'Status', 'Nº Chamado', 'Solicitante', 'Ocorrência', 'Título', 'Analista','Data Encerramento', 'Remover Filtro', command=drop_selecao_busca)
    drop_busca.config(font=fonte_padrao, bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF", highlightthickness=0, relief=RIDGE, width=15, cursor="hand2")
    drop_busca['menu'].insert_separator(7)
    drop_busca.grid(row=0, column=7, padx=(0,2))
    clique_busca.set('Filtrar por...')
    ent_busca = Entry(frame2, width=30, font=fonte_padrao, justify='center')
    ent_busca.grid(row=0, column=8, ipady=4)
    ent_busca.bind('<Return>', pesquisar_bind)

    def muda_busca(e):
        image_busca = Image.open('imagens\\lupa_over.png')
        resize_busca = image_busca.resize((25, 25))
        nova_image_busca = ImageTk.PhotoImage(resize_busca)
        btn_busca.photo = nova_image_busca
        btn_busca.config(image=nova_image_busca, fg='#7c7c7c')
    def volta_busca(e):
        image_busca = Image.open('imagens\\lupa.png')
        resize_busca = image_busca.resize((25, 25))
        nova_image_busca = ImageTk.PhotoImage(resize_busca)
        btn_busca.photo = nova_image_busca
        btn_busca.config(image=nova_image_busca, fg='#ffffff')
    image_busca = Image.open('imagens\\lupa.png')
    resize_busca = image_busca.resize((25, 25))
    nova_image_busca = ImageTk.PhotoImage(resize_busca)
    btn_busca = Button(frame2, image=nova_image_busca, bg="#1d366c", fg='#FFFFFF', command=pesquisar,
                            borderwidth=0, relief=RIDGE, activebackground="#1d366c", activeforeground="#7c7c7c", cursor="hand2")
    btn_busca.photo = nova_image_busca
    btn_busca.grid(row=0, column=9, padx=(4,30))
    btn_busca.bind("<Enter>", muda_busca)
    btn_busca.bind("<Leave>", volta_busca)


    image_usuario = Image.open('imagens\\usuario.png')
    resize_usuario = image_usuario.resize((30, 30))
    nova_image_usuario = ImageTk.PhotoImage(resize_usuario)
    lbluserlogado = Label(frame2, image=nova_image_usuario, text=usuariologado, compound="left", font=fonte_padrao, bg="#1d366c", fg='#fff000')
    lbluserlogado.photo = nova_image_ferramentas
    lbluserlogado.grid(row=0, column=10)


    frame2.grid_columnconfigure(6, weight=1)
    frame2.grid_columnconfigure(6, weight=1)

    # -------------------FRAME3-------------------#
    style = ttk.Style()
    # style.theme_use('default')
    style.configure('Treeview',
                    background='#ffffff',
                    rowheight=24,
                    fieldbackground='#ffffff',
                    font=fonte_padrao)
    style.configure("Treeview.Heading",
                    foreground='#000000',
                    background="#ffffff",
                    font=fonte_padrao)
    style.map('Treeview', background=[('selected', '#1d366c')])

    tree_principal = ttk.Treeview(frame3, selectmode='browse')
    vsb = ttk.Scrollbar(frame3, orient="vertical", command=tree_principal.yview)
    vsb.pack(side=RIGHT, fill='y')
    tree_principal.configure(yscrollcommand=vsb.set)
    vsbx = ttk.Scrollbar(frame3, orient="horizontal", command=tree_principal.xview)
    vsbx.pack(side=BOTTOM, fill='x')
    tree_principal.configure(xscrollcommand=vsbx.set)
    tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
    tree_principal["columns"] = ("1", "2", "3", "4", "5", "6", "7", "8", "9")
    tree_principal['show'] = 'headings'
    tree_principal.column("1", width=88, anchor='c')
    tree_principal.column("2", width=116, anchor='c')
    tree_principal.column("3", width=180, anchor='c')
    tree_principal.column("4", width=140, anchor='c')
    tree_principal.column("5", width=140, anchor='c')
    tree_principal.column("6", width=300, anchor='c')
    tree_principal.column("7", width=140, anchor='c')
    tree_principal.column("8", width=120, anchor='c')
    tree_principal.column("9", width=160, anchor='c')
    tree_principal.heading("1", text="Nº Chamado")
    tree_principal.heading("2", text="Data de Abertura")
    tree_principal.heading("3", text="Solicitante")
    tree_principal.heading("4", text="Tipo")
    tree_principal.heading("5", text="Ocorrência")
    tree_principal.heading("6", text="Título")
    tree_principal.heading("7", text="Analista")
    tree_principal.heading("8", text="Status")
    tree_principal.heading("9", text="Data de Encerramento")
    tree_principal.tag_configure('par', background='#e9e9e9')
    tree_principal.tag_configure('impar', background='#ffffff')
    tree_principal.bind("<Double-1>", duploclique_tree_principal)
    frame3.grid_columnconfigure(0, weight=1)
    frame3.grid_columnconfigure(3, weight=1)

    # -------------------FRAME4-------------------#
    lbllacuna = Label(frame4, text='', bg="#1d366c")
    lbllacuna.grid(row=0, column=1)

    frame4.grid_columnconfigure(0, weight=1)
    frame4.grid_columnconfigure(3, weight=1)
    contador()
    root.bind("<F1>",abrirchamado_bind)
    root.bind("<F2>",atendimento_bind)
    root.bind("<F3>",visualizar_chamado_bind)
    root.bind("<F4>",ferramentas_bind)
    atualizar_lista_principal()

    root.title(titulo_todos)
    root.state('zoomed')
    root.iconbitmap('imagens\\ico.ico')

    ###-------------------- VERIFICADOR DE VERSÃO DO SISTEMA----------------------
    verifica_versao = cursor.execute("SELECT * FROM dbo.versao")
    versao_banco = verifica_versao.fetchone()
    if versao_banco[0] != versao:
        messagebox.showerror('Atualização do Sistema:',
                             f'Versão do Sistema desatualizada.\n\nVersão(local) do Sistema: {versao}\nVersão atualizada do Sistema: {versao_banco[0]}\n\nReinicie sua máquina para que o software se atualize automaticamente ou entre em contato com o TI.',
                             parent=root)
        root.destroy()
    root.mainloop()

#/////////////////////////////FIM PRINCIPAL/////////////////////////////

# /////////////////////////////INICIO LOGIN/////////////////////////////
def login():
    splash_root.destroy()
    global root2
    root2 = Tk()
    root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
    root2.unbind_class("Button", "<Key-space>")
    root2.focus_force()
    root2.grab_set()

    def sair():
        root2.destroy()

    def entrar():
        user = euser.get()
        senha = esenha.get()
        r = cursor.execute("SELECT * FROM dbo.analista WHERE login=?", (user,))
        result = r.fetchone()
        if result != None:
            clique.set("Analista")

        if user == "" or senha == "":
            messagebox.showwarning('Login: Erro', 'Digite o Usuário ou Senha.', parent=root2)
        else:
            if clique.get() == "Usuário":
                server_name = '192.168.1.19'
                domain_name = 'gvdobrasil'
                server = Server(server_name, get_info=ALL)
                try:
                    Connection(server, user='{}\\{}'.format(domain_name, user), password=senha, authentication=NTLM,
                               auto_bind=True)
                    global nivel_acesso
                    nivel_acesso = 0
                    global usuariologado
                    usuariologado = user
                    principal()
                except:
                    messagebox.showwarning('Login: Erro', 'Usuário ou senha inválidos.', parent=root2)
            else:
                r = cursor.execute("SELECT * FROM dbo.analista WHERE login=?", (user,))
                result = r.fetchone()
                if result is None:
                    messagebox.showwarning('Login: Erro', 'Usuário ou Senha inválidos.', parent=root2)
                else:
                    r = cursor.execute("SELECT * FROM dbo.analista WHERE login=?", (user,))
                    for login in r.fetchall():
                        filtro_user = login[1]
                        filtro_pwd = login[3]
                    if bcrypt.checkpw(senha.encode("utf-8"), filtro_pwd.encode("utf-8")):
                        nivel_acesso = 1
                        usuariologado = login[2]
                        contador()
                        principal()
                    else:
                        messagebox.showwarning('Login: Erro', 'Usuário ou Senha inválidos.', parent=root2)

    def entrar_bind(event):
        entrar()

    frame0 = Frame(root2, bg='#ffffff')
    frame0.grid(row=0, column=0, stick='nsew')
    root2.grid_rowconfigure(0, weight=1)
    root2.grid_columnconfigure(0, weight=1)
    frame1 = Frame(frame0, bg="#1d366c")
    frame1.pack(side=TOP, fill=X, expand=False, anchor='center')
    frame2 = Frame(frame0, bg='#ffffff')
    frame2.pack(side=TOP, fill=X, expand=False, anchor='center', pady=10)
    frame3 = Frame(frame0, bg='#ffffff')
    frame3.pack(side=TOP, fill=X, expand=False, anchor='center')
    frame4 = Frame(frame0, bg='#ffffff')
    frame4.pack(side=TOP, fill=X, expand=False, anchor='center', pady=10)
    frame5 = Frame(frame0, bg='#1d366c')
    frame5.pack(side=TOP, fill=X, expand=False, anchor='center')

    image_login = Image.open('imagens\\login.png')
    resize_login = image_login.resize((35, 35))
    nova_image_login = ImageTk.PhotoImage(resize_login)

    lbllogin = Label(frame1, image=nova_image_login, text=" Login", compound="left", bg='#1d366c',
          fg='#FFFFFF', font=fonte_titulos)
    lbllogin.photo = nova_image_login
    lbllogin.grid(row=0, column=1)
    frame1.grid_columnconfigure(0, weight=1)
    frame1.grid_columnconfigure(2, weight=1)

    Label(frame2, text="Modo de Acesso:", bg='#ffffff', fg='#000000', font=fonte_padrao).grid(row=0, column=1,
                                                                                              sticky="w")
    clique = StringVar()
    clique.set("Usuário")
    drop = OptionMenu(frame2, clique, "Usuário", "Analista")
    drop.config(bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c', activeforeground="#FFFFFF",
                highlightthickness=0, relief=RIDGE, width=9, font=fonte_padrao, cursor="hand2")
    drop.grid(row=0, column=2, pady=10)
    frame2.grid_columnconfigure(0, weight=1)
    frame2.grid_columnconfigure(3, weight=1)

    Label(frame3, text="Usuário:", bg='#ffffff', fg='#000000', font=fonte_padrao).grid(row=1, column=1, sticky="w")
    euser = Entry(frame3, width=30, font=fonte_padrao)
    euser.grid(row=1, column=2, sticky="w", padx=5, pady=10)
    euser.focus_force()
    euser.bind("<Return>", entrar_bind)
    Label(frame3, text="Senha:", font=fonte_padrao, bg='#ffffff', fg='#000000').grid(row=2, column=1, sticky="w")
    esenha = Entry(frame3, show="*", width=30, font=fonte_padrao)
    esenha.grid(row=2, column=2, sticky="w", padx=5, pady=10)
    esenha.bind("<Return>", entrar_bind)
    frame3.grid_columnconfigure(0, weight=1)
    frame3.grid_columnconfigure(3, weight=1)

    bt1 = Button(frame4, text='Entrar', bg='#1d366c', fg='#FFFFFF', activebackground='#1d366c',
                 activeforeground="#FFFFFF", highlightthickness=0, width=10, relief=RIDGE, command=entrar,
                 font=fonte_padrao, cursor="hand2")
    bt1.grid(row=0, column=1, pady=5, padx=5)

    bt2 = Button(frame4, text='Sair', width=10, relief=RIDGE, command=sair, font=fonte_padrao, cursor="hand2")
    bt2.grid(row=0, column=2, pady=5, padx=5)
    frame4.grid_columnconfigure(0, weight=1)
    frame4.grid_columnconfigure(3, weight=1)

    Label(frame5, text="", bg='#1d366c', fg='#FFFFFF', font=fonte_padrao).grid(row=0, column=1, sticky="ew")
    frame5.grid_columnconfigure(0, weight=1)
    frame5.grid_columnconfigure(2, weight=1)
    '''root2.update()
    largura = frame0.winfo_width()
    altura = frame0.winfo_height()
    print(largura, altura)'''
    window_width = 330
    window_height = 275
    screen_width = root2.winfo_screenwidth()
    screen_height = root2.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    #root2.resizable(0, 0)
    root2.iconbitmap('imagens\\ico.ico')
    root2.title(titulo_todos)
# /////////////////////////////FIM LOGIN/////////////////////////////


#/////////////////////////////ROOT SPLASH/////////////////////////////
splash_root =Tk()
frame0 = Frame(splash_root, bg="#000000")
frame0.grid(row=0, column=0, stick='nsew')
splash_root.grid_rowconfigure(0, weight=1)
splash_root.grid_columnconfigure(0, weight=1)

image_splash = Image.open('imagens\\splash.jpg')
resize_splash = image_splash.resize((572, 152))
nova_image_splash = ImageTk.PhotoImage(resize_splash)
lbl_splash = Label(frame0, image=nova_image_splash)
lbl_splash.pack()
lbl_splash.photo = nova_image_splash

#splash_root.update()
#largura = splash_root.winfo_width()
#altura = splash_root.winfo_height()
#print(largura, altura)
window_width = 572
window_height = 152
screen_width = splash_root.winfo_screenwidth()
screen_height = splash_root.winfo_screenheight()
x_cordinate = int((screen_width / 2) - (window_width / 2))
y_cordinate = int((screen_height / 2) - (window_height / 2))
splash_root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
splash_root.resizable(0, 0)
splash_root.overrideredirect(True)
splash_root.after(1000, login)
#/////////////////////////////FIM ROOT SPLASH/////////////////////////////

#/////////////////////////////BANCO DE DADOS/////////////////////////////

# Banco de Dados

'''
server = 'tcp:GVBRSRV01,1433'
database = 'helpdesk'
username = 'sa'
password = 'senha'
conectar = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = conectar.cursor()
'''
'''
server = 'GVBRSRV01\SQLEXPRESS2014'
database = 'helpdesk'
conectar = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + '; Trusted_Connection=yes')
cursor = conectar.cursor()
'''
#/////////////////////////////FIM BANCO DE DADOS/////////////////////////////
server = '192.168.1.19\SQLEXPRESS2014'
database = 'helpdesk'
#username = 'sa'
#password = 'gv2K20ADM'
username = 'acesso_rede'
password = 'senha'
conectar = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = conectar.cursor()

mainloop()