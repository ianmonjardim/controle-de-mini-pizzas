import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
from datetime import datetime

def update_stock(flavor, quantity, operation):
    # Carregar ou criar o arquivo Excel
    try:
        wb = load_workbook('controle_pizza.xlsx')
        sheet = wb.active
    except FileNotFoundError:
        wb = Workbook()
        sheet = wb.active
        sheet.append(['Date', 'Operation', 'Flavor', 'Quantity', 'Stock'])

    # Atualizar o estoque
    current_stock = 0
    for row in sheet.iter_rows(min_row=2, max_col=4):
        if row[2].value.lower() == flavor:
            if operation == 'Entrada':
                current_stock += row[3].value
            elif operation == 'Saída':
                current_stock -= row[3].value

    new_stock = current_stock + quantity if operation == 'Entrada' else current_stock - quantity
    sheet.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), operation, flavor, quantity, new_stock])

    # Salvar o arquivo Excel
    wb.save('controle_pizza.xlsx')

def register_entry():
    flavor = entry_flavor.get().strip().lower()
    if flavor not in ['presunto e queijo', 'calabresa']:
        messagebox.showerror('Erro', 'Sabor inválido. Por favor, escolha entre "Presunto e queijo" ou "Calabresa".')
        return
    quantity = int(entry_quantity.get())
    update_stock(flavor, quantity, 'Entrada')
    messagebox.showinfo('Sucesso', f'{quantity} mini pizzas de {flavor} registradas como entrada no estoque.')

def register_exit():
    flavor = entry_flavor.get().strip().lower()
    if flavor not in ['presunto e queijo', 'calabresa']:
        messagebox.showerror('Erro', 'Sabor inválido. Por favor, escolha entre "Presunto e queijo" ou "Calabresa".')
        return
    quantity = int(entry_quantity.get())
    update_stock(flavor, quantity, 'Saída')
    messagebox.showinfo('Sucesso', f'{quantity} mini pizzas de {flavor} registradas como saída do estoque.')

def clear_stock():
    confirmation = messagebox.askyesno('Confirmação', 'Tem certeza de que deseja limpar todo o estoque?')
    if confirmation:
        wb = Workbook()
        sheet = wb.active
        sheet.append(['Date', 'Operation', 'Flavor', 'Quantity', 'Stock'])
        wb.save('controle_pizza.xlsx')
        messagebox.showinfo('Sucesso', 'Todo o estoque foi apagado.')

def remove_item():
    flavor = entry_flavor.get().strip().lower()
    confirmation = messagebox.askyesno('Confirmação', f'Tem certeza de que deseja apagar todas as entradas/saídas de {flavor}?')
    if confirmation:
        wb = load_workbook('controle_pizza.xlsx')
        sheet = wb.active
        rows_to_delete = []
        for row in sheet.iter_rows(min_row=2, max_col=4):
            if row[2].value.lower() == flavor:
                rows_to_delete.append(row)
        for row in rows_to_delete:
            sheet.delete_rows(row[0].row)
        wb.save('controle_pizza.xlsx')
        messagebox.showinfo('Sucesso', f'Todas as entradas/saídas de {flavor} foram apagadas.')

# Configuração da janela principal
root = tk.Tk()
root.title('Contole de Pizzas')

# Componentes da interface
label_flavor = tk.Label(root, text='Sabor da Pizza:')
label_flavor.grid(row=0, column=0, padx=10, pady=5)

entry_flavor = tk.Entry(root)
entry_flavor.grid(row=0, column=1, padx=10, pady=5)

label_quantity = tk.Label(root, text='Quantidade:')
label_quantity.grid(row=1, column=0, padx=10, pady=5)

entry_quantity = tk.Entry(root)
entry_quantity.grid(row=1, column=1, padx=10, pady=5)

button_entry = tk.Button(root, text='Registrar Entrada', command=register_entry)
button_entry.grid(row=2, column=0, padx=10, pady=5)

button_exit = tk.Button(root, text='Registrar Saída', command=register_exit)
button_exit.grid(row=2, column=1, padx=10, pady=5)

button_clear = tk.Button(root, text='Limpar Estoque', command=clear_stock)
button_clear.grid(row=3, column=0, padx=10, pady=5)

button_remove = tk.Button(root, text='Remover Item', command=remove_item)
button_remove.grid(row=3, column=1, padx=10, pady=5)

# Loop principal da interface
root.mainloop()
