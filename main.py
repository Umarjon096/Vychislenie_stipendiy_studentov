from itog_generator import *


# libraries Import
from tkinter import Tk, filedialog, messagebox
import customtkinter
import subprocess

# Main Window Properties

window = Tk()
window.title("Sborka otchet")
window.geometry("800x280")
window.configure(bg="#a2a7b3")


# Функция выбора папки для Entry_id1
def browse_input_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        Entry_id1.delete(0, 'end')
        Entry_id1.insert(0, folder_selected)

# Функция выбора папки для Entry_id6
def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        Entry_id6.delete(0, 'end')
        Entry_id6.insert(0, folder_selected)


# Обработка изменения чекбокса
def toggle_checkbox():
    if Checkbox_id8.get() == 1:
        Entry_id6.configure(state="disabled")
        Button_id7.configure(state="disabled")
    else:
        Entry_id6.configure(state="normal")
        Button_id7.configure(state="normal")


def start():
    input_path = Entry_id1.get()
    output_path = Entry_id6.get()
    checkbox = Checkbox_id8.get() == 1
    exlporer = Checkbox_explorer.get() == 1
    xlsx = Checkbox_xlsx.get() == 1
    output_file = 'itog.xlsx'
    output = input_path + '/' + output_file

    if not input_path:
        messagebox.showerror("Xatolik", "Birinchi popka ko'rsatilmagan.")
        return
    if not os.path.isdir(input_path):
        messagebox.showerror("Xatolik", "Birinchi popka noto'g'ri yoki mavjud emas.")
        return

    if not checkbox:
        if not output_path:
            messagebox.showerror("Xatolik", "Ikkinchi popka ko'rsatilmagan. Yoki galochka qo'ying.")
            return
        if not os.path.isdir(output_path):
            messagebox.showerror("Xatolik", "Ikkinchi popka noto'g'ri yoki mavjud emas.")
            return
    else:
        output_path = input_path

    try:
        os.remove(output)
    except FileNotFoundError:
        print("❗ Файл не найден — удалять нечего. Продолжим")
    except PermissionError:
        print("❗ Файл занят (возможно, открыт в Excel). Закрой файл и попробуй снова.")
        messagebox.showerror("Error", "❗ Файл занят (возможно, открыт в Excel). Закрой файл и попробуй снова.")
        exit()
    except Exception as e:
        print(f"❗ Неизвестная ошибка при удалении: {e}")
        messagebox.showerror("Error", f"❗ Неизвестная ошибка при удалении выходного файла:\n {e}")
        exit()

    generator(input_path, output_path, output_file)
    messagebox.showinfo("Success", "Fayllar ochildi, bitta faylga saqlandi")
    # Открыть проводник и выделить файл
    if exlporer: subprocess.run(f'explorer /select,"{output.replace('/', '\\')}"')
    if xlsx: os.startfile(output)



Entry_id1 = customtkinter.CTkEntry(
    master=window,
    placeholder_text="popka",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=550,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    )
Entry_id1.place(x=10, y=50)
Button_id3 = customtkinter.CTkButton(
    master=window,
    text="Obzor",
    font=("undefined", 16),
    text_color="#000000",
    hover=True,
    hover_color="#b7b3b3",
    height=30,
    width=95,
    border_width=2,
    corner_radius=10,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    command=browse_input_folder,
    )
Button_id3.place(x=580, y=50)
Label_id5 = customtkinter.CTkLabel(
    master=window,
    text="Fayl saqlash uchun popkani ko'rsating",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=250,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    )
Label_id5.place(x=10, y=100)
Button_id7 = customtkinter.CTkButton(
    master=window,
    text="Obzor",
    font=("undefined", 16),
    text_color="#000000",
    hover=True,
    hover_color="#a8a4a4",
    height=30,
    width=95,
    border_width=2,
    corner_radius=10,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    command=browse_output_folder,
    )
Button_id7.place(x=580, y=140)
Label_id2 = customtkinter.CTkLabel(
    master=window,
    text="Fayllar turgan popkani ko'rsating",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=250,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    )
Label_id2.place(x=10, y=10)
Button_id4 = customtkinter.CTkButton(
    master=window,
    text="Start",
    font=("undefined", 26),
    text_color="#000000",
    hover=True,
    hover_color="#b2aeae",
    height=50,
    width=150,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    command=start,
    )
Button_id4.place(x=630, y=210)

Checkbox_id8 = customtkinter.CTkCheckBox(
    master=window,
    text="Fayllar turgan popkaga",
    text_color="#000000",
    border_color="#000000",
    fg_color="#808080",
    hover_color="#808080",
    corner_radius=4,
    border_width=2,
command=toggle_checkbox,
    )
Checkbox_id8.place(x=310, y=100)



Checkbox_explorer = customtkinter.CTkCheckBox(
    master=window,
    text="Faylni provodnikda ko'rsatish",
    text_color="#000000",
    border_color="#000000",
    fg_color="#808080",
    hover_color="#808080",
    corner_radius=4,
    border_width=2,
    )
Checkbox_explorer.place(x=110, y=200)




Checkbox_xlsx = customtkinter.CTkCheckBox(
    master=window,
    text="Faylni ochish (Excel)",
    text_color="#000000",
    border_color="#000000",
    fg_color="#808080",
    hover_color="#808080",
    corner_radius=4,
    border_width=2,
    )
Checkbox_xlsx.place(x=310, y=200)


Entry_id6 = customtkinter.CTkEntry(
    master=window,
    placeholder_text="popka",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=550,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    )
Entry_id6.place(x=10, y=140)

#run the main loop
window.mainloop()

