from tkinter import *
import tkinter.ttk as ttk
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
from openpyxl import load_workbook
import os

class PhoneBook():
    def __init__(self):
        self.imenik = tk.Tk()
        self.imenik.title("Telefonski imenik")
        self.ime = StringVar()
        self.prezime = StringVar()
        self.telefon = StringVar()
        self.nadji = StringVar()
        self.nadji.set("Pretrazi...")
        self.rbutton = IntVar()
        self.lista = []
        self.lista1 = []

        unosPodataka = tk.LabelFrame(self.imenik, text="Unos podataka", padx=5, pady=5)
        imeLabel = tk.Label(unosPodataka, text="Ime: ")
        prezimeLabel = tk.Label(unosPodataka, text="Prezime: ")
        telefonLabel = tk.Label(unosPodataka, text="Telefon: ")
        self.imeEntry = tk.Entry(unosPodataka, textvariable=self.ime)
        self.prezimeEntry = tk.Entry(unosPodataka, textvariable=self.prezime)
        self.telefonEntry = tk.Entry(unosPodataka, textvariable=self.telefon)
        Dodaj = tk.Button(unosPodataka, text="Dodaj", command=self.dodaj)

        unosPodataka.grid(row=0, column=0)

        imeLabel.grid(row=0, column=0)
        self.imeEntry.grid(row=0, column=1)
        prezimeLabel.grid(row=1, column=0)
        self.prezimeEntry.grid(row=1, column=1)
        telefonLabel.grid(row=2, column=0)
        self.telefonEntry.grid(row=2, column=1)
        Dodaj.grid(row=0, column=3, columnspan=3, rowspan=3, sticky=W + E + N + S)

        alatke = tk.LabelFrame(self.imenik, text="Alatke", padx=5, pady=5)
        alatke.grid(row=0, column=1)
        obrisiSve = tk.Button(alatke, text="Obrisi sve", command=self.obrisiSVE)
        obrisiSve.grid(row=0, column=0, sticky=W + E)
        izmeni = tk.Button(alatke, text="Izmeni", command=self.izmeni)
        izmeni.grid(row=0, column=1, sticky=W + E)
        obrisiSelektovano = tk.Button(alatke, text="Obrisi selektovano", command=self.ObrisiSelektovano)
        obrisiSelektovano.grid(row=1, column=0, sticky=W + E)
        memorisi = tk.Button(alatke, text="Memorisi", command=self.memorisi)
        memorisi.grid(row=1, column=1, sticky=W + E)
        ispisiBazu = tk.Button(alatke, text="Ispisi bazu", command=self.ispisiBazu)
        ispisiBazu.grid(row=2, column=1, sticky=W + E)
        obrisiBazu = tk.Button(alatke, text="Obrisi bazu", command=self.obrisiBazu)
        obrisiBazu.grid(row=2, column=0, sticky=W + E)

        pretragaKontakata = tk.LabelFrame(self.imenik, text="Pretraga kontakata", padx=5, pady=5)
        pretragaKontakata.grid(row=2, column=0, columnspan=2)
        self.pretraga = tk.Entry(pretragaKontakata, textvariable=self.nadji).pack(side=LEFT)
        pretrazi = tk.Button(pretragaKontakata, text="Pretrazi", command=self.pretrazi).pack(side=LEFT)
        poImenu = tk.Radiobutton(pretragaKontakata, text="Po imenu", variable=self.rbutton, value=1).pack(side=LEFT)
        poPrezimenu = tk.Radiobutton(pretragaKontakata, text="Po prezimenu", variable=self.rbutton, value=2).pack(side=LEFT)
        poBrojuTelefona = tk.Radiobutton(pretragaKontakata, text="Po broju telefona", variable=self.rbutton, value=3)
        poBrojuTelefona.pack(side=LEFT)

        self.spisak = ttk.Treeview(self.imenik)
        self.spisak['show'] = 'headings'
        self.spisak["columns"] = ("ime", "prezime", "telefon")
        self.spisak.column("ime", width=230, minwidth=230)
        self.spisak.column("prezime", width=230, minwidth=230)
        self.spisak.column("telefon", width=140, minwidth=140, stretch=tk.NO)
        self.spisak.heading("ime", text="Ime")
        self.spisak.heading("prezime", text="Prezime")
        self.spisak.heading("telefon", text="Telefon")
        self.spisak.grid(row=1, column=0, columnspan=2)

        self.imeEntry.focus()
        self.imenik.mainloop()


    def dodaj(self):
        if self.validacija(self.ime,self.prezime,self.telefon,self.imeEntry,self.prezimeEntry,self.telefonEntry)==True:
            self.spisak.insert("", "end", values=((self.ime.get()), (self.prezime.get()), (self.telefon.get())))
            self.lista.append((self.ime.get(), self.prezime.get(), self.telefon.get()))
            self.ime.set("")
            self.prezime.set("")
            self.telefon.set("")
            self.imeEntry.focus()


    def obrisiSVE(self):
        if self.lista != []:
            self.spisak.delete(*self.spisak.get_children())
            self.lista.clear()
        else:
            messagebox.showwarning("Obavestenje","Lista je vec prazna.")


    def validacija(self,ime,prezime,telefon,imeEntry,prezimeEntry,telefonEntry):
        if ime.get()=="":
            messagebox.showerror("Greska", "Ime je obavezan podatak. Unesite ime.")
            imeEntry.focus()
            return False
        elif prezime.get()=="":
            messagebox.showerror("Greska","Prezime je obavezan podatak. Unesite prezime.")
            prezimeEntry.focus()
            return False
        elif telefon.get()=="":
            messagebox.showerror("Greska","Broj telefona je obavezan podatak. Unesite broj telefona.")
            telefonEntry.focus()
            return False
        elif not telefon.get().isnumeric():
            messagebox.showerror("Greska","Broj telefona moze sadrzati samo numericke karaktere(brojeve).")
            telefon.set("")
            telefonEntry.focus()
            return False
        elif telefon.get()[:3]!="381":
            messagebox.showerror("Neispravan broj telefona","Broj telefona mora pocinjati sa 381")
            telefon.set("")
            telefonEntry.focus()
            return False
        elif telefon.get()[3]=="0":
            messagebox.showerror("Neispravan broj telefona","Broj telefona nije odgovarajuceg formata.")
            telefon.set("")
            telefonEntry.focus()
            return False
        elif len(telefon.get())<11 or len(telefon.get())>12:
            messagebox.showerror("Neispravan broj telefona","Neispravna duzina broja telefona.")
            telefon.set("")
            telefonEntry.focus()
            return False
        else:
            return True


    def ObrisiSelektovano(self):
        if self.lista != [] and self.spisak.selection() != ():
            pitanje = messagebox.askquestion("Brisanje","Da li zelite da obrisete ovaj kontakt iz baze?",icon="question")
            if pitanje == "yes":
                potvrda = False
                br_obrisanih = 0
                br_pokusaja = 0
                pristup_bazi = 0
                for selektovan in self.spisak.selection():
                    ime = str(tuple(self.spisak.item(selektovan)['values'])[0])
                    prezime = str(tuple(self.spisak.item(selektovan)['values'])[1])
                    telefon = str(tuple(self.spisak.item(selektovan)['values'])[2])
                    for tupl in range(len(self.lista)):
                        if self.lista[tupl] == (ime,prezime,telefon):
                            self.lista.pop(tupl)
                            break
                    try:
                        kontakti = load_workbook(filename="BazaKontakata.xlsx")
                        kontakti.active = 0
                        Sheet = kontakti.active
                        broj = 2
                        for value in Sheet.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
                            if value[0] == ime and value[1] == prezime and str(value[2]) == str(telefon):
                                Sheet.delete_rows(idx=broj, amount=1)
                                kontakti.save(filename="BazaKontakata.xlsx")
                                potvrda = True
                                br_obrisanih+=1
                            else:
                                broj+=1
                    except:
                        if pristup_bazi == 0:
                            messagebox.showinfo("Obavestenje", "Ne postoji baza sa kontaktima.")
                            pristup_bazi += 1
                    br_pokusaja +=1
                    self.spisak.delete(selektovan)

                if potvrda == True:
                    if br_obrisanih == 1:
                        messagebox.showinfo("Obavestenje", "Oznaceni kontakt je uspesno obrisan iz baze.")
                    elif br_obrisanih > 1:
                        messagebox.showinfo("Obavestenje", "Oznaceni kontakti su uspesno obrisani iz baze.")
                else:
                    if br_pokusaja == 1:
                        messagebox.showinfo("Obavestenje", "Kontakt koji zelite obrisati ne postoji u bazi kontakata, ali ce biti obrisan sa liste.")
                    elif br_pokusaja > 1:
                        messagebox.showinfo("Obavestenje", "Kontakti koje zelite obrisati ne postoje u bazi kontakata, ali ce biti obrisani sa liste.")
            else:
                for selektovan in self.spisak.selection():
                    ime = str(tuple(self.spisak.item(selektovan)['values'])[0])
                    prezime = str(tuple(self.spisak.item(selektovan)['values'])[1])
                    telefon = str(tuple(self.spisak.item(selektovan)['values'])[2])
                    for tupl in range(len(self.lista)):
                        if self.lista[tupl] == (ime,prezime,telefon):
                            self.lista.pop(tupl)
                            break
                    self.spisak.delete(selektovan)
                messagebox.showinfo("Obavestenje","Kontakt je prividno obrisan sa liste, ali se i dalje nalazi u bazi.")
        else:
            messagebox.showwarning("Obavestenje","Niste oznacili kontakte koje zelite da obrisete.")


    def memorisi(self):
        if self.lista != []:
            try:
                kontakti = load_workbook(filename="BazaKontakata.xlsx")
            except:
                kreiraj = Workbook()
                kreiraj.save(filename="BazaKontakata.xlsx")
                kontakti = load_workbook(filename="BazaKontakata.xlsx")

            kontakti.active = 0
            Sheet = kontakti.active

            Sheet.column_dimensions['A'].width = 20.0
            Sheet.column_dimensions['B'].width = 20.0
            Sheet.column_dimensions['C'].width = 20.0

            if Sheet["A1"].value == "" or Sheet["A1"].value == None:
                Sheet["A1"] = "IME"
                Sheet["B1"] = "PREZIME"
                Sheet["C1"] = "TELEFON"

            for kontakt in self.lista:
                ime = kontakt[0]
                prezime = kontakt[1]
                telefon = kontakt[2]

                broj = 2
                for value in Sheet.iter_rows(min_row=1, min_col=1, max_col=3, values_only=True):
                    if Sheet["A" + str(broj)].value == "" or Sheet["A" + str(broj)].value == None:
                        Sheet["A" + str(broj)] = ime
                        Sheet["B" + str(broj)] = prezime
                        Sheet["C" + str(broj)] = telefon
                        br = 0
                        for value in Sheet.iter_rows(min_row=broj+1, min_col=1, max_col=3, values_only=True):
                            if Sheet["A" + str(broj+1+br)].value == ime and Sheet["B" + str(broj+1+br)].value == prezime and Sheet["C" + str(broj+1+br)].value == telefon:
                                Sheet["A" + str(broj + 1 + br)] = ""
                                Sheet["B" + str(broj + 1 + br)] = ""
                                Sheet["C" + str(broj + 1 + br)] = ""
                            br+=1
                    elif Sheet["A" + str(broj)].value == ime and Sheet["B" + str(broj)].value == prezime and Sheet["C" + str(broj)].value == telefon:
                        break
                    else:
                        broj+=1

            kontakti.save(filename="BazaKontakata.xlsx")
            messagebox.showinfo("Obavestenje","Uspesno ste snimili kontakte u bazu.")
        else:
            messagebox.showwarning("Obavestenje","Niste dodali kontakte u listu da biste ih memorisali u bazu.")


    def pretrazi(self):
        for kontakt in self.lista:
            if self.rbutton.get()==1 and self.nadji.get() in kontakt[0]:
                self.lista1.append(kontakt)
            elif self.rbutton.get()==2 and self.nadji.get() in kontakt[1]:
                self.lista1.append(kontakt)
            elif self.rbutton.get()==3 and self.nadji.get() in kontakt[2]:
                self.lista1.append(kontakt)

        if self.lista1==[]:
            messagebox.showwarning("Obavestenje","Trazeni kontakt nije pronadjen.")
            self.nadji.set("Pretrazi...")
        else:
            self.spisak.delete(*self.spisak.get_children())
            self.lista.clear()
            for kontakt in self.lista1:
                self.spisak.insert("", "end", values=((kontakt[0]), (kontakt[1]), (kontakt[2])))
                self.lista.append((kontakt[0], kontakt[1], kontakt[2]))
            self.lista1.clear()
            self.nadji.set("Pretrazi...")


    def izmeni(self):
        if len(self.spisak.selection()) == 1:
            self.izmenaKontakta = tk.Toplevel()
            self.izmenaKontakta.title("Izmena")
            self.kontakt_za_izmenu = self.spisak.selection()[0]
            self.listaIzmena = list(self.spisak.item(self.kontakt_za_izmenu).values())[2]
            self.staroIme = self.listaIzmena[0]
            self.staroPrezime = self.listaIzmena[1]
            self.stariTelefon = str(self.listaIzmena[2])
            self.novoIme = StringVar()
            self.novoPrezime = StringVar()
            self.novBrojTelefona = StringVar()

            tk.Label(self.izmenaKontakta, text="Unesite novo ime:").grid(row=0, column=0)
            tk.Label(self.izmenaKontakta, text="Unesite novo prezime:").grid(row=1, column=0)
            tk.Label(self.izmenaKontakta, text="Unesite nov broj telefona:").grid(row=2, column=0)
            tk.Button(self.izmenaKontakta, text="Izmeni", command=self.izmeni_kontakt).grid(row=3, column=0, columnspan=2, sticky=W+E)
            self.novoImeEntry = tk.Entry(self.izmenaKontakta, textvariable=self.novoIme)
            self.novoPrezimeEntry = tk.Entry(self.izmenaKontakta, textvariable=self.novoPrezime)
            self.novBrojTelefonaEntry = tk.Entry(self.izmenaKontakta, textvariable=self.novBrojTelefona)
            self.novoImeEntry.grid(row=0, column=1)
            self.novoPrezimeEntry.grid(row=1, column=1)
            self.novBrojTelefonaEntry.grid(row=2, column=1)

            self.novoImeEntry.insert(0, self.listaIzmena[0])
            self.novoPrezimeEntry.insert(0, self.listaIzmena[1])
            self.novBrojTelefonaEntry.insert(0, str(self.listaIzmena[2]))

        elif len(self.spisak.selection()) < 1:
            messagebox.showwarning("Obavestenje","Niste selektovali kontakt koji zelite da izmenite.")
        else:
            messagebox.showwarning("Obavestenje","Ne mozete izmeniti vise kontakata istovremeno. Selektujte samo jedan kontakt.")


    def izmeni_kontakt(self):
        if self.validacija(self.novoIme,self.novoPrezime,self.novBrojTelefona,self.novoImeEntry,self.novoPrezimeEntry,self.novBrojTelefonaEntry) == True:
            id = 0
            for kontakt in self.lista:
                if kontakt == (self.staroIme,self.staroPrezime,self.stariTelefon):
                    self.lista.pop(id)
                    self.lista.insert(id,(self.novoIme.get(),self.novoPrezime.get(),self.novBrojTelefona.get()))
                    break
                else:
                    id += 1
            self.spisak.delete(*self.spisak.get_children())
            for kontakt in self.lista:
                self.spisak.insert("", "end", values=((kontakt[0]), (kontakt[1]), (kontakt[2])))

            self.izmenaKontakta.destroy()

            pitanje = messagebox.askquestion("Pitanje", "Da li zelite da izmenite ovaj kontakt i u bazi?", icon="question")
            if pitanje == "yes":
                try:
                    kontakti = load_workbook(filename="BazaKontakata.xlsx")
                    kontakti.active = 0
                    Sheet = kontakti.active
                    promena = False
                    broj = 2
                    for value in Sheet.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
                        if value[0] == self.staroIme and value[1] == self.staroPrezime and str(value[2]) == str(self.stariTelefon):
                            Sheet["A" + str(broj)] = self.novoIme.get()
                            Sheet["B" + str(broj)] = self.novoPrezime.get()
                            Sheet["C" + str(broj)] = self.novBrojTelefona.get()
                            kontakti.save(filename="BazaKontakata.xlsx")
                            promena = True
                        else:
                            broj += 1

                    if promena == True:
                        messagebox.showinfo("Obavestenje", "Zeljena promena kontakta je izvrsena u bazi kontakata.")
                    else:
                        messagebox.showinfo("Obavestenje", "Kontakt koji zelite promeniti se ne nalazi u bazi kontakata.")

                except:
                    messagebox.showinfo("Obavestenje","Ne postoji baza sa kontaktima.")
            else:
                messagebox.showinfo("Obavestenje","Kontakt je promenjen na listi kontakata, ali ne i u bazi.")


    def ispisiBazu(self):
        try:
            kontakti = load_workbook(filename="BazaKontakata.xlsx")
            kontakti.active = 0
            Sheet = kontakti.active

            for value in Sheet.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
                if value[0]!=None and value[1]!=None and value[2]!=None:
                    self.spisak.insert("", "end", values=((value[0]), (value[1]), (value[2])))
                    self.lista.append((value[0],value[1],value[2]))
                else:
                    pass
            if self.lista == []:
                messagebox.showwarning("Obavestenje","Baza kontakata je prazna.")
        except:
            messagebox.showinfo("Obavestenje","Ne postoji baza sa kontaktima.")


    def obrisiBazu(self):
        upozorenje = messagebox.askquestion("Upozorenje","Upozorenje! Bice obrisana baza i svi kontakti u njoj. Da li ste sigurni da zelite da nastavite?",icon="warning")
        if upozorenje == "yes":
            try:
                filePath = os.path.join(os.getcwd(), "BazaKontakata.xlsx")
                os.remove(filePath)
                messagebox.showinfo("Obavestenje", "Baza sa kontaktima je uspesno obrisana.")
            except:
                messagebox.showinfo("Obavestenje", "Ne postoji baza sa kontaktima.")
        else:
            pass


start = PhoneBook()