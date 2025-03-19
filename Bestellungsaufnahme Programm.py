import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import pathlib
from tkcalendar import DateEntry
from datetime import datetime

class MyApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Bestellungsaufnahme Programm")
        self.window.geometry("700x400+300+200")
        self.window.resizable(False, False)
        self.window.configure(bg="#2D2B55")
        self.window.protocol("WM_DELETE_WINDOW", self.quitFunction)

        self.excelFont = openpyxl.styles.Font(name="Arial", size=16, bold=True)

        self.file = pathlib.Path("Bestellungs Liste.xlsx")
        if self.file.exists():
            pass
        else:
            self.file = openpyxl.Workbook()
            self.sheet = self.file.active
            self.sheet["A1"] = "Name"
            self.sheet["A1"].font = self.excelFont
            self.sheet["B1"] = "Versandart"
            self.sheet["B1"].font = self.excelFont
            self.sheet["C1"] = "Produkt"
            self.sheet["C1"].font = self.excelFont
            self.sheet["D1"] = "Datum"
            self.sheet["D1"].font = self.excelFont
            self.sheet["E1"] = "Uhrzeit"
            self.sheet["E1"].font = self.excelFont
            self.sheet["F1"] = "Preis"
            self.sheet["F1"].font = self.excelFont
            self.sheet["G1"] = "Stückzahl"
            self.sheet["G1"].font = self.excelFont
            self.sheet["H1"] = "Summe"
            self.sheet["H1"].font = self.excelFont
            self.sheet["I1"] = "Beschreibung"
            self.sheet["I1"].font = self.excelFont

            self.file.save("Bestellungs Liste.xlsx")

        tk.Label(self.window, text="Bitte das Formular ausfüllen:", bg="#2D2B55", fg="white", font="Arial 12").place(x=20, y=20)

        tk.Label(self.window, text="Name", bg="#2D2B55", fg="white", font="Arial 11").place(x=50, y=100)
        tk.Label(self.window, text="Versandart", bg="#2D2B55", fg="white", font="Arial 11").place(x=50, y=150)
        tk.Label(self.window, text="Produkt", bg="#2D2B55", fg="white", font="Arial 11").place(x=360, y=150)
        tk.Label(self.window, text="Preis", bg="#2D2B55", fg="white", font="Arial 11").place(x=50, y=200)
        tk.Label(self.window, text="Stückzahl", bg="#2D2B55", fg="white", font="Arial 11").place(x=360, y=200)
        tk.Label(self.window, text="Datum", bg="#2D2B55", fg="white", font="Arial 11").place(x=50, y=250)
        tk.Label(self.window, text="Uhrzeit", bg="#2D2B55", fg="white", font="Arial 11").place(x=360, y=250)
        tk.Label(self.window, text="Beschreibung", bg="#2D2B55", fg="white", font="Arial 11").place(x=50, y=300)


        self.nameEntry = tk.Entry(self.window)
        self.nameEntry.place(x=190, y=100, width=420)

        self.sendungCombo = ttk.Combobox(self.window, value=["Abholung", "Versand"])
        self.sendungCombo.set("Abholung")
        self.sendungCombo.place(x=190, y=150, width=150)

        self.produktCombo = ttk.Combobox(self.window, value=["Produkt 1", "Produkt 2", "Produkt 3", "Produkt 4", "Produkt 5"])
        self.produktCombo.set("Produkt 1")
        self.produktCombo.place(x=460, y=150, width=150)

        self.price = tk.Entry(self.window)
        self.price.place(x=190, y=200, width=150)

        self.mengeEntry = tk.Entry(self.window)
        self.mengeEntry.place(x=460, y=200, width=150)

        self.date_picker = DateEntry(self.window, width=12, bg="#2D2B55", fg="white", borderwidth=2, locale="de_DE")
        self.date_picker.place(x=190, y=250, width=150)

        self.hour_combo = ttk.Combobox(self.window, values=[str(i).zfill(2) for i in range(24)], width=3)
        self.hour_combo.place(x=460, y=250, width=75)
        self.hour_combo.set(datetime.now().strftime("%H"))

        self.minute_combo = ttk.Combobox(self.window, values=[str(i).zfill(2) for i in range(60)], width=3)
        self.minute_combo.place(x=540, y=250, width=72)
        self.minute_combo.set(datetime.now().strftime("%M"))

        self.description_textbox = tk.Text(self.window)
        self.description_textbox.place(x=190, y=300, width=420, height=35)

        self.submitButton = tk.Button(self.window, text="Ergänzen", bg="#2D2B55", fg="white", command=self.submitFunction)
        self.submitButton.place(x=190, y=350, height=40, width=100)

        self.quitButton = tk.Button(self.window, text="Verlassen", bg="#2D2B55", fg="white", command=self.quitFunction)
        self.quitButton.place(x=510, y=350, height=40, width=100)

    def submitFunction(self):
        name = self.nameEntry.get()
        sendung = self.sendungCombo.get()
        product = self.produktCombo.get()
        price = self.price.get()
        menge = self.mengeEntry.get()
        math = float(price) * float(menge)
        summe = str(math) + "€"
        date = self.date_picker.get()
        time = f"{self.hour_combo.get()}:{self.minute_combo.get()}"
        description = self.description_textbox.get("1.0", tk.END)

        self.file = openpyxl.load_workbook("Bestellungs Liste.xlsx")
        self.sheet = self.file.active

        for cell in self.sheet[1]:
            cell.font = self.excelFont

        self.sheet.append([name, sendung, product, date, time, price, menge, summe, description])

        self.file.save(r'Bestellungs Liste.xlsx')

        self.nameEntry.delete(0, tk.END)
        self.sendungCombo.set("Abholung")
        self.produktCombo.set("Produkt 1")
        self.price.delete(0, tk.END)
        self.mengeEntry.delete(0, tk.END)
        self.description_textbox.delete("1.0", tk.END)

    def helpFunction(self):
        msg = messagebox.showinfo("Hinweis", "Die Menge Bitte mit der Einheit angeben", icon="warning")

    def quitFunction(self):
        msg = messagebox.askquestion("Verlassen", "Sind sie sich sicher?", icon="warning")
        if msg == "yes":
            self.window.quit()

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = MyApp()
    app.run()
