import sys
import threading
import queue
import time
import os
import builtins
import ctypes
import customtkinter as ctk

from PIL import Image
from pathlib import Path


# ============================================================
# IMPORTA OS SCRIPTS DO DOSSIÊ
# ============================================================

import dossie_01_preparar_base
import dossie_02_montar_final
import dossie_03_aplicar_revisao


# ============================================================
# IDENTIDADE DO APLICATIVO NO WINDOWS
# ============================================================
# Ajuda o Windows a tratar o programa como um aplicativo próprio
# e não apenas como "Python".

if sys.platform == "win32":
    try:
        APP_ID = "Portfolio.GeradorDossies.1.0"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


# ============================================================
# CONFIGURAÇÃO DO TEMA
# ============================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ============================================================
# CAMINHOS DOS ARQUIVOS
# ============================================================

def get_resource_path(relative_path):
    """
    Encontra corretamente os arquivos do aplicativo.

    Funciona:
    - rodando com python app.py
    - rodando com pythonw.exe
    - rodando futuramente como .exe com PyInstaller
    """

    try:
        # Quando o programa estiver empacotado com PyInstaller
        base_path = Path(sys._MEIPASS)

    except AttributeError:
        # Quando estiver executando diretamente pelo Python
        base_path = Path(__file__).resolve().parent

    return base_path / relative_path


# Pasta das imagens
IMG_DIR = get_resource_path("IMAGENS")

# Ícone do aplicativo
ICON_PATH = IMG_DIR / "icone_app.ico"


# ============================================================
# APLICATIVO
# ============================================================

class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        # ====================================================
        # CONFIGURAÇÕES DA JANELA
        # ====================================================

        self.title("Gerador de Dossiês")
        self.geometry("750x550")
        self.resizable(False, False)

        # Aplica o ícone depois que a janela terminar de carregar.
        # Isso evita que o CustomTkinter sobrescreva o ícone.
        self.after(200, self.aplicar_icone)

        # Fila utilizada para substituir o input() dos scripts
        self.input_queue = queue.Queue()

        # Configuração principal do layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ====================================================
        # LADO ESQUERDO - MASCOTE
        # ====================================================

        self.frame_mascote = ctk.CTkFrame(
            self,
            width=250,
            corner_radius=0,
            fg_color="transparent"
        )

        self.frame_mascote.grid(
            row=0,
            column=0,
            rowspan=2,
            sticky="nsew",
            padx=10,
            pady=20
        )

        self.frame_mascote.grid_rowconfigure(1, weight=1)

        # ====================================================
        # CARREGA AS IMAGENS DOS MASCOTES
        # ====================================================

        self.img_esperando = self.load_mascot_image(
            "Feliz esperando.png"
        )

        self.img_trabalhando = self.load_mascot_image(
            "Analisando.png"
        )

        self.img_sucesso = self.load_mascot_image(
            "Comemorando.png"
        )

        self.img_erro = self.load_mascot_image(
            "Duvida.png"
        )

        # Mascote inicial
        self.mascote_label = ctk.CTkLabel(
            self.frame_mascote,
            text="",
            image=self.img_esperando
        )

        self.mascote_label.grid(
            row=1,
            column=0,
            sticky="nsew"
        )

        # ====================================================
        # LADO DIREITO
        # ====================================================

        self.frame_main = ctk.CTkFrame(
            self,
            corner_radius=15
        )

        self.frame_main.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=20,
            pady=20
        )

        self.frame_main.grid_rowconfigure(1, weight=1)
        self.frame_main.grid_columnconfigure(0, weight=1)

        # ====================================================
        # TÍTULO
        # ====================================================

        self.lbl_title = ctk.CTkLabel(
            self.frame_main,
            text="Automação do Dossiê",
            font=ctk.CTkFont(
                size=24,
                weight="bold"
            )
        )

        self.lbl_title.grid(
            row=0,
            column=0,
            pady=(15, 5),
            sticky="n"
        )

        # ====================================================
        # CONSOLE / BALÃO DE FALA
        # ====================================================

        self.console_box = ctk.CTkTextbox(
            self.frame_main,
            wrap="word",
            corner_radius=10,
            font=ctk.CTkFont(size=12)
        )

        self.console_box.grid(
            row=1,
            column=0,
            padx=15,
            pady=10,
            sticky="nsew"
        )

        self.console_box.insert(
            "0.0",
            "Olá! Sou o seu assistente virtual do Programa. "
            "Qual relatório vamos gerar hoje?\n\n"
        )

        self.console_box.configure(
            state="disabled"
        )

        # ====================================================
        # CAMPO PARA DIGITAR RESPOSTAS
        # ====================================================

        self.frame_input = ctk.CTkFrame(
            self.frame_main,
            fg_color="transparent"
        )

        self.frame_input.grid(
            row=2,
            column=0,
            padx=15,
            pady=(0, 15),
            sticky="ew"
        )

        self.frame_input.grid_columnconfigure(
            0,
            weight=1
        )

        self.entry_input = ctk.CTkEntry(
            self.frame_input,
            placeholder_text="Digite sua resposta aqui..."
        )

        self.entry_input.grid(
            row=0,
            column=0,
            padx=(0, 10),
            sticky="ew"
        )

        self.entry_input.bind(
            "<Return>",
            self.send_input
        )

        self.btn_send = ctk.CTkButton(
            self.frame_input,
            text="Enviar",
            width=80,
            command=self.send_input
        )

        self.btn_send.grid(
            row=0,
            column=1
        )

        # ====================================================
        # RODAPÉ - BOTÕES
        # ====================================================

        self.frame_buttons = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )

        self.frame_buttons.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=20,
            pady=(0, 20)
        )

        self.frame_buttons.grid_columnconfigure(
            (0, 1, 2),
            weight=1
        )

        # ====================================================
        # BOTÃO - FASE 1
        # ====================================================

        self.btn_fase1 = ctk.CTkButton(
            self.frame_buttons,
            text="1. Criação Base",
            fg_color="#1F4E5F",
            hover_color="#2B6F89",
            command=lambda: self.run_script(1)
        )

        self.btn_fase1.grid(
            row=0,
            column=0,
            padx=5,
            sticky="ew"
        )

        # ====================================================
        # BOTÃO - FASE DE REVISÃO
        # ====================================================

        self.btn_fase3 = ctk.CTkButton(
            self.frame_buttons,
            text="2. Gerar Revisão",
            fg_color="#4B7F52",
            hover_color="#5D9B65",
            command=lambda: self.run_script(3)
        )

        self.btn_fase3.grid(
            row=0,
            column=1,
            padx=5,
            sticky="ew"
        )

        # ====================================================
        # BOTÃO - FINAL DIRETO
        # ====================================================

        self.btn_fase2 = ctk.CTkButton(
            self.frame_buttons,
            text="Gerar Final Direto",
            fg_color="#8A3C3C",
            hover_color="#A84A4A",
            command=lambda: self.run_script(2)
        )

        self.btn_fase2.grid(
            row=0,
            column=2,
            padx=5,
            sticky="ew"
        )

        # ====================================================
        # REDIRECIONAMENTO DE INPUT / PRINT
        # ====================================================

        self.setup_io_redirection()


    # ========================================================
    # APLICA O ÍCONE DO CAPACETE
    # ========================================================

    def aplicar_icone(self):

        try:
            if ICON_PATH.exists():
                self.iconbitmap(str(ICON_PATH))

        except Exception as e:
            print(
                f"Não foi possível carregar o ícone: {e}"
            )


    # ========================================================
    # CARREGA AS IMAGENS DO MASCOTE
    # ========================================================

    def load_mascot_image(self, file_name):

        path = IMG_DIR / file_name

        if path.exists():

            image = Image.open(path)

            return ctk.CTkImage(
                light_image=image,
                dark_image=image,
                size=(220, 220)
            )

        return None


    # ========================================================
    # TROCA O MASCOTE
    # ========================================================

    def change_mascot(self, state):

        imagens = {

            "esperando":
                self.img_esperando,

            "trabalhando":
                self.img_trabalhando,

            "sucesso":
                self.img_sucesso,

            "erro":
                self.img_erro
        }

        img = imagens.get(
            state,
            self.img_esperando
        )

        if img:

            self.after(
                0,
                lambda: self.mascote_label.configure(
                    image=img
                )
            )


    # ========================================================
    # REDIRECIONA PRINT E INPUT DOS SCRIPTS
    # ========================================================

    def setup_io_redirection(self):

        class StdoutRedirector:

            def __init__(self, app):
                self.app = app

            def write(self, text):

                if text:

                    self.app.after(
                        0,
                        self.app.append_to_console,
                        text
                    )

            def flush(self):
                pass


        # Tudo que seria mostrado no terminal
        # passa a aparecer dentro da caixa do aplicativo.
        sys.stdout = StdoutRedirector(self)
        sys.stderr = StdoutRedirector(self)


        # ====================================================
        # SUBSTITUI O INPUT() DO PYTHON
        # ====================================================

        def gui_input(prompt=""):

            if prompt:

                self.after(
                    0,
                    self.append_to_console,
                    str(prompt)
                )

            # O script pausa aqui até você responder
            # no campo do aplicativo e clicar em Enviar.
            return self.input_queue.get()


        builtins.input = gui_input


    # ========================================================
    # ESCREVE MENSAGEM NO CONSOLE
    # ========================================================

    def append_to_console(self, text):

        self.console_box.configure(
            state="normal"
        )

        self.console_box.insert(
            "end",
            text
        )

        self.console_box.see(
            "end"
        )

        self.console_box.configure(
            state="disabled"
        )


    # ========================================================
    # ENVIA A RESPOSTA DIGITADA
    # ========================================================

    def send_input(self, event=None):

        text = self.entry_input.get().strip()

        if text:

            self.append_to_console(
                f"{text}\n"
            )

            self.input_queue.put(
                text
            )

            self.entry_input.delete(
                0,
                "end"
            )


    # ========================================================
    # ATIVA OU DESATIVA OS BOTÕES
    # ========================================================

    def toggle_buttons(self, state):

        estado = "normal" if state else "disabled"

        self.btn_fase1.configure(
            state=estado
        )

        self.btn_fase2.configure(
            state=estado
        )

        self.btn_fase3.configure(
            state=estado
        )


    # ========================================================
    # INICIA UMA FASE
    # ========================================================

    def run_script(self, fase):

        self.toggle_buttons(False)

        self.change_mascot(
            "trabalhando"
        )

        thread = threading.Thread(
            target=self._execute_fase,
            args=(fase,),
            daemon=True
        )

        thread.start()


    # ========================================================
    # EXECUTA OS SCRIPTS
    # ========================================================

    def _execute_fase(self, fase):

        try:

            sys.argv = [
                f"dossie_{fase}.py"
            ]

            # =================================================
            # FASE 1
            # =================================================

            if fase == 1:

                self.append_to_console(
                    "\n"
                    "> Iniciando Fase 1: Preparação da Base...\n"
                    "⏳ Aguarde, estou analisando as pastas e "
                    "criando o Word base. "
                    "Isso pode demorar um pouquinho...\n\n"
                )

                dossie_01_preparar_base.main()


            # =================================================
            # FINAL DIRETO
            # =================================================

            elif fase == 2:

                self.append_to_console(
                    "\n"
                    "> Iniciando Geração Final Direta...\n"
                    "⏳ Aguarde, estou montando "
                    "os PDFs finais...\n\n"
                )

                dossie_02_montar_final.main()


            # =================================================
            # REVISÃO
            # =================================================

            elif fase == 3:

                self.append_to_console(
                    "\n"
                    "> Iniciando Fase de Revisão...\n"
                    "⏳ Aguarde, estou mesclando os anexos "
                    "e gerando os arquivos finais...\n\n"
                )

                dossie_03_aplicar_revisao.main()


            # =================================================
            # SUCESSO
            # =================================================

            self.append_to_console(
                "\n"
                "✅ Processo concluído com sucesso! "
                "Tudo pronto.\n"
            )

            self.change_mascot(
                "sucesso"
            )


        except SystemExit:

            self.append_to_console(
                "\n✅ Processo finalizado.\n"
            )

            self.change_mascot(
                "sucesso"
            )


        except Exception as e:

            self.append_to_console(
                "\n"
                f"❌ Ops! Ocorreu um erro: {e}\n"
            )

            self.change_mascot(
                "erro"
            )


        finally:

            time.sleep(4)

            self.after(
                0,
                lambda: self.toggle_buttons(True)
            )

            self.change_mascot(
                "esperando"
            )


# ============================================================
# INICIA O APLICATIVO
# ============================================================

if __name__ == "__main__":

    app = App()

    app.mainloop()