from typing import Optional, Tuple, Union
import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl, xlrd
import pathlib
from openpyxl import Workbook, workbook
from tkcalendar import Calendar, DateEntry

# Setando a aparencia do sistema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()
        self.apperence()
        self.todo_sistema()

    def layout_config(self):
        self.title("Sistema de Gestão de Clientes - Nutricionista Karina Macaco")
        self.geometry("700x500")

    def apperence(self):
        self.lb_apm = ctk.CTkLabel(
            self, text="Tema", bg_color="transparent", text_color=["#000", "#fff"]
        ).place(x=50, y=440)
        self.opt_apm = ctk.CTkOptionMenu(
            self, values=["Dark", "Light", "System"], command=self.change_apm
        ).place(x=50, y=465)

    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)

    def todo_sistema(self):
        frame = ctk.CTkFrame(
            self,
            width=700,
            height=50,
            corner_radius=0,
            bg_color="teal",
            fg_color="teal",
        ).place(x=0, y=10)
        title = ctk.CTkLabel(
            frame,
            text="Sistema de Gestão de Clientes",
            font=("Century Gothic", 24),
            text_color="#fff",
            bg_color="teal",
        ).place(x=170, y=20)
        span = ctk.CTkLabel(
            self,
            text="Por Favor, preencha todos os campos do formulário!",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        ).place(x=50, y=70)

        ficheiro = pathlib.Path("Clientes.xlsx")

        if ficheiro.exists():
            pass
        else:
            ficheiro = Workbook()
            folha = ficheiro.active
            folha["A1"] = "Nome Completo"
            folha["B1"] = "Telefone"
            folha["C1"] = "Idade"
            folha["D1"] = "Genero"
            folha["E1"] = "Endereço"
            folha["F1"] = "Email"
            folha["G1"] = "Observações"

            ficheiro.save("Clientes.xlsx")

        # Funções
        def submit():
            # pegando os dados dos entrys
            name = name_value.get()
            phone = phone_value.get()
            age = calendar_date
            age = age_value.get()
            gender = gender_combobox.get()
            adress = adress_value.get()
            email = email_value.get()
            obs = obs_entry.get(0.0, END)

            if (
                name == "" or phone == "" or email == ""
            ):  ## Criar validação para vazio e existente.
                messagebox.showerror(
                    "Sistema",
                    "Operação não concluida!\nPor favor Preencha todos os campos",
                )
            else:
                ficheiro = openpyxl.load_workbook("Clientes.xlsx")
                folha = ficheiro.active
                folha.cell(column=1, row=folha.max_row + 1, value=name)
                folha.cell(column=2, row=folha.max_row, value=phone)
                folha.cell(column=3, row=folha.max_row, value=age)
                folha.cell(column=4, row=folha.max_row, value=gender)
                folha.cell(column=5, row=folha.max_row, value=adress)
                folha.cell(column=6, row=folha.max_row, value=email)
                folha.cell(column=7, row=folha.max_row, value=obs)

                ficheiro.save(r"Clientes.xlsx")
                messagebox.showinfo("Sistema", "Dados Salvos com Sucesso!")

        def clear():
            name_value.set("")
            phone_value.set("")
            age_value.set("")
            adress_value.set("")
            email_value.set("")
            obs_entry.delete(0.0, END)

            # test variables

        name_value = StringVar()
        phone_value = StringVar()
        age_value = StringVar()
        adress_value = StringVar()
        email_value = StringVar()

        # Entrys
        name_entry = ctk.CTkEntry(
            self,
            textvariable=name_value,
            width=350,
            font=("Century Gothic", 16),
            fg_color="transparent",
        )
        contact_entry = ctk.CTkEntry(
            self,
            textvariable=phone_value,
            width=200,
            font=("Century Gothic", 16),
            fg_color="transparent",
        )
        age_entry = ctk.CTkEntry(
            self,
            textvariable=age_value,
            width=150,
            font=("Century Gothic", 16),
            fg_color="transparent",
        )
        adress_entry = ctk.CTkEntry(
            self,
            textvariable=adress_value,
            width=200,
            font=("Century Gothic", 16),
            fg_color="transparent",
        )
        email_entry = ctk.CTkEntry(
            self,
            textvariable=email_value,
            width=200,
            font=("Century Gothic", 16),
            fg_color="transparent",
        )
        calendar_date = DateEntry(
            self,
            selectmode="day",
            year=2023,
            width=12,
            font=("Century Gothic", 12),
            background="teal",
            foreground="white",
            borderwidth=3,
        )
        # Combo box
        gender_combobox = ctk.CTkComboBox(
            self,
            values=["Masculino", "Feminino", "Outro"],
            font=("Century Gothic", 14),
        )
        gender_combobox.set("Feminino")

        # Entrada de observações
        obs_entry = ctk.CTkTextbox(
            self,
            width=600,
            height=75,
            font=("arial", 12),
            border_color="#aaa",
            border_width=2,
            fg_color="transparent",
        )

        # Botoes submit e limpar
        button_submit = ctk.CTkButton(
            self,
            text="Salvar dados".upper(),
            command=submit,
            fg_color="#151",
            hover_color="#131",
        ).place(x=375, y=465)
        button_submit = ctk.CTkButton(
            self,
            text="Limpar Campos".upper(),
            command=clear,
            fg_color="#555",
            hover_color="#333",
        ).place(x=525, y=465)

        # Labels
        lb_name = ctk.CTkLabel(
            self,
            text="Nome Completo:",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        )
        lb_contact = ctk.CTkLabel(
            self,
            text="Telefone para contato:",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        )
        lb_age = ctk.CTkLabel(
            self,
            text="Data de nascimento:",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        )
        lb_gender = ctk.CTkLabel(
            self,
            text="Gênero:",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        )
        lb_adress = ctk.CTkLabel(
            self,
            text="Endereço:",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        )
        lb_email = ctk.CTkLabel(
            self,
            text="Email:",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        )
        lb_obs = ctk.CTkLabel(
            self,
            text="Observações:",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        )

        # Posicionando na tela
        lb_name.place(x=50, y=120)
        name_entry.place(x=50, y=150)

        lb_contact.place(x=450, y=120)
        contact_entry.place(x=450, y=150)

        lb_age.place(x=300, y=190)
        age_entry.place(x=300, y=220)
        # calendar_date.place(x=300, y=220)

        lb_gender.place(x=500, y=190)
        gender_combobox.place(x=500, y=220)

        lb_adress.place(x=50, y=260)  # 260
        adress_entry.place(x=50, y=290)  # 290

        lb_email.place(x=50, y=190)  # 190
        email_entry.place(x=50, y=220)  # 220

        lb_obs.place(x=50, y=330)
        obs_entry.place(x=50, y=360)


if __name__ == "__main__":
    app = App()
    app.mainloop()