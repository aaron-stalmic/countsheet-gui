import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import font
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import time
import requests

TITLE = "Countsheet Updater"
CURRENT_WAREHOUSE = None


class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        """
        Use our completion list as our drop down selection menu, arrows move
        through menu.
        """
        self._completion_list = sorted(completion_list, key=str.lower)
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = self._completion_list  # Setup our popup menu.

    def autocomplete(self, delta=0):
        """
        Autocomplete the Combobox, delta may be 0/1/-1 to cycle through
        possible hits.
        """
        # Need to delete selection otherwise we would fix current position.
        if delta:
            self.delete(self.position, tk.END)
        else:
            self.position = len(self.get())
        _hits = []
        for element in self._completion_list:
            if element.lower().startswith(self.get().lower()):
                _hits.append(element)
        # If we have a new hit list, keep this in mind.
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        # Only allow cycling if we are in a known hit list.
        if _hits == self._hits and self._hits:
            self._hit_index = (self._hit_index + delta) % len(self._hits)
        # Perform the autocompletion
        if self._hits:
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        """
        Event handler for the keyrelease event on this widget.
        """
        if event.keysym == 'BackSpace':
            self.delete(self.index(tk.INSERT), tk.END)
            self.position = self.index(tk.END)
        if event.keysym == 'Left':
            if self.position < self.index(tk.END):  # Delete the selection.
                self.delete(self.position, tk.END)
            # else
            #    self.position = self.position - 1  # Delete one character.
            #    self.delete(self.position, tk.END)
        if event.keysym == 'Right' or event.keysym == 'KP_Enter':
            self.position == self.index(tk.END)  # Go to end (no selection)
        if len(event.keysym) == 1:
            self.autocomplete()


class WHSheet(gspread.Worksheet):
    def __init__(self, sheet, worksheet):
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
                       'credentials.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open(sheet)
        self.ws = sh.worksheet(worksheet)

    def getCol(self, col):
        collist = self.ws.col_values(col)
        length = len(collist) - 1
        while collist[length] == '':
            length -= 1
        return collist[0:length+1]

    def setValue(self, r, c, value):
        self.ws.update_cell(r, c, value)

    def getValue(self, r, c):
        return self.ws.cell(r, c).input_value

    def addRow(self, values):
        self.ws.append_row(values)


class SelectWindow(tk.Frame):
    def __init__(self, parent, item_list, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.window = tk.Toplevel(self.parent)
        self.window.grab_set()
        self.window.focus()
        self.window.geometry('450x650')
        self.window.update()
        self.item_list = sorted(item_list)

        self.scrollbar = tk.Scrollbar(self.window, width=100)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(self.window, yscrollcommand=self.scrollbar.set, width=35)
        for item in self.item_list:
            self.listbox.insert(tk.END, item)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH)

        self.scrollbar.config(command=self.listbox.yview)

        self.listbox.bind('<<ListboxSelect>>', self.immediately)

    def immediately(self, e):
        self.parent.item.set(self.item_list[self.listbox.curselection()[0]])


class MainApplication(tk.Frame):
    def __init__(self, parent, #inventory, sheet, history,
                 *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        #self.inventory = inventory
        #self.sheet = sheet
        #self.history = history
        # Make a menu bar
        if not RESTRICTED:
            self.menubar = tk.Menu(self.parent)
            menu = tk.Menu(self.menubar, tearoff=0)
            self.menubar.add_cascade(label="Change Warehouse", menu=menu)
            menu.add_command(label="Townsend", command=self.townsend)
            menu.add_command(label="Lakeland", command=self.lakeland)
            self.parent.config(menu=self.menubar)
        # Get the item list
        self.item_list = []
        for i in range(1, 4):
            self.item_list += self.get_list(i)
        tk.Label(self.parent, text="Item:").grid(row=0, column=0)
        self.item = tk.StringVar()
        self._item = AutocompleteCombobox(self.parent, textvariable=self.item,
                                          values=self.item_list)
        self._item.set_completion_list(self.item_list)
        self._item.grid(row=0, column=1, columnspan=3, sticky=tk.W+tk.E)
        self.action = tk.IntVar()
        tk.Radiobutton(self.parent, text="Add",
                       variable=self.action, value=1,
                       indicatoron=0, width=10).grid(row=1, column=0,
                                                     sticky=tk.W+tk.S)
        tk.Radiobutton(self.parent, text="Remove",
                       variable=self.action, value=-1,
                       indicatoron=0, width=10).grid(row=2, column=0,
                                                     sticky=tk.W+tk.N)
        tk.Label(self.parent, text="Amount:").grid(row=1, column=1,
                                                   sticky=tk.E)
        self.amount = tk.StringVar()
        self._amount = tk.Entry(self.parent, textvariable=self.amount, width=8)
        self._amount.grid(row=1, column=2, sticky=tk.W)
        tk.Label(self.parent, text="Reason:").grid(row=2, column=1,
                                                   sticky=tk.E)
        self.reason = tk.Text(self.parent, width=25, height=4, wrap=tk.WORD)
        self.reason.grid(row=2, column=2, columnspan=2, sticky=tk.W)
        tk.Button(self.parent, text="Select",
                  command=self.select).grid(row=1, column=3, sticky=tk.E)
        tk.Button(self.parent, text="Submit",
                  command=self.submit).grid(row=3, column=0)
        width, height = self.parent.grid_size()
        for x in range(1, width):
            tk.Grid.columnconfigure(self.parent, x, weight=1)
        for y in range(height):
            tk.Grid.rowconfigure(self.parent, y, weight=1)

    def get_list(self, c):
        """
        Gets a list of items in a column.
        """
        global CURRENT_WAREHOUSE
        if CURRENT_WAREHOUSE == 'Townsend':
            self.inventory = WHSheet('Townsend Warehouse Inventory Sheet', 'Inventory')
        else:
            self.inventory = WHSheet('Lakeland Warehouse Inventory Sheet', 'Inventory')
        return self.inventory.getCol(c)

    def select(self):
        SelectWindow(self, self.item_list)

    def submit(self):
        global CURRENT_WAREHOUSE
        try:
            amount = int(self.amount.get())
            if amount < 0:
                raise ValueError
        except ValueError:
            if self.amount.get() == "":
                pass
            else:
                error = "Amount must be a positive integer."
                tk.messagebox.showerror(TITLE, error)
        else:
            item = self.item.get()
            if item not in self.item_list:
                tk.messagebox.showerror(TITLE, "Invalid item selected.")
            elif not self.action.get():
                tk.messagebox.showerror(TITLE, "Please select an action.")
            else:
                if self.action.get() == 1:
                    op = '+'
                else:
                    op = '-'
                columns = [1, 5, 10]
                try:
                    if CURRENT_WAREHOUSE == 'Townsend':
                        self.sheet = WHSheet('Townsend Warehouse Inventory Sheet', 'Townsend Count Sheet')
                        self.history = WHSheet('Townsend Warehouse Inventory Sheet', 'History')
                    else:
                        self.sheet = WHSheet('Lakeland Warehouse Inventory Sheet', 'Lakeland Count Sheet')
                        self.history = WHSheet('Lakeland Warehouse Inventory Sheet', 'History')
                    print(CURRENT_WAREHOUSE)
                    for c in columns:
                        col = self.sheet.getCol(c)[1:]
                        for r in range(len(col)):
                            if col[r] == item:
                                current = self.sheet.getValue(r+2, c+1)
                                self.sheet.setValue(r+2, c+1,
                                                    current + op + str(amount))
                                time = datetime.datetime.now()
                                time = time.strftime('%m/%d/%y %I:%M %p')
                                self.history.addRow([item,
                                                    amount * self.action.get(),
                                                    time,
                                                    self.reason.get('1.0',
                                                                    'end-1c')])
                                tk.messagebox.showinfo(TITLE, "Successfully updated.")
                                return
                    tk.messagebox.showerror(TITLE, "Item not in countsheet.")
                except requests.ConnectionError:
                    tk.messagebox.showerror(TITLE, "Failed to add changes! Check connection.")

    def townsend(self):
        global CURRENT_WAREHOUSE
        root.wm_title(TITLE + ' - Townsend')
        CURRENT_WAREHOUSE = 'Townsend'
        self.destroy()
        self.__init__(root) #, inventory, townsend, history)


    def lakeland(self):
        global CURRENT_WAREHOUSE
        root.wm_title(TITLE + ' - Lakeland')
        CURRENT_WAREHOUSE = 'Lakeland'
        self.destroy()
        self.__init__(root) #, l_inventory, lakeland, l_history)


if __name__ == '__main__':
    RESTRICTED = False

    root = tk.Tk()
    default_font = font.nametofont('TkDefaultFont')
    default_font.configure(size=14)
    root.option_add('*Font', default_font)
    root.geometry('450x250')
    if RESTRICTED:
        root.wm_title(TITLE + ' - Lakeland')
        CURRENT_WAREHOUSE = 'Lakeland'
    else:
        root.wm_title(TITLE + ' - Townsend')
        CURRENT_WAREHOUSE = 'Townsend'
    MainApplication(root)
    root.mainloop()
