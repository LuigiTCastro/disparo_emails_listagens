import tkinter as tk
from tkinter import messagebox, simpledialog
from functools import partial


# Construção dos meses
dict_months = {
    'JANEIRO' : '01',
    'FEVEREIRO' : '02',
    'MARÇO' : '03',
    'ABRIL' : '04',
    'MAIO' : '05',
    'JUNHO' : '06',
    'JULHO' : '07',
    'AGOSTO' : '08',
    'SETEMBRO' : '09',
    'OUTUBRO' : '10',
    'NOVEMBRO' : '11',
    'DEZEMBRO' : '12',
}
    
# --------------------------------------------------------------------------------------------------

# INPUT BOX
def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

def get_month():
    root = tk.Tk()
    root.deiconify()
    center_window(root)
    root.withdraw()
    input_value = simpledialog.askstring('Input', 'Digite o mês de referência (ex.: JANEIRO): ').upper()
    if input_value is None:
        return
    if not input_value in dict_months:
        msg = messagebox.showinfo('Erro', 'Nome inválido.')
        get_month()
    root.destroy()
    return input_value

def get_option():
    root = tk.Tk()
    root.deiconify()
    center_window(root)
    root.withdraw()
    input_value = simpledialog.askstring('Input', "Se deseja enviar para os e-mails de teste, digite 'TESTE', caso contrário, apenas aperte em 'OK'").upper()
    root.destroy()
    return input_value
