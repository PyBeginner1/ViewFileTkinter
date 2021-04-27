import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

root=tk.Tk()
root.geometry('500x500')
root.title("ExcelView")
root.pack_propagate(False)          #Dont resize. pack_propagate(0) tells tkinter to let the parent control its own size, rather than letting it's size be determined by the children.
root.resizable(0,0)                 #cant resize


#create a frame treeview
frame=tk.LabelFrame(root, text="Excel Data")
frame.place(height=250, width=500)

#frame for open file dialog
file_frame=tk.LabelFrame(root, text="Open File")
file_frame.place(height=100, width=400, rely=0.55, relx=0)



#create buttons
button1=tk.Button(file_frame, text="Browse", command=lambda:browse_file())
button1.place(relx=0.5, rely=0.6)

button2=tk.Button(file_frame, text="Load File", command=lambda:load_file())
button2.place(rely=0.6, relx=0.3)

label_file=ttk.Label(file_frame, text="No File Selected")
label_file.place(relx=0, rely=0)

#create treeview
my_tree=ttk.Treeview(frame)
my_tree.place(relheight=1, relwidth=1)      #it says to take up the whole frame

#create scrollbar for x & y axis
tree_scrollx=tk.Scrollbar(frame, orient='horizontal', command=my_tree.xview)

tree_scrolly=tk.Scrollbar(frame, orient="vertical", command=my_tree.yview)

#configure scroll
my_tree.config(xscrollcommand=tree_scrollx.set, yscrollcommand=tree_scrolly.set)
tree_scrollx.pack(side="bottom", fill="x")
tree_scrolly.pack(side="right",fill="y")

def browse_file():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file["text"] = filename
    return None

def load_file():
    file_path=label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)


    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    my_tree["column"] = list(df.columns)
    my_tree["show"] = "headings"
    for column in my_tree["column"]:
        my_tree.heading(column,text=column)     #col heading is column name

    #rows
    df_rows=df.to_numpy().tolist()     #turns data frame into list of lists
    for row in df_rows:
        my_tree.insert("","end",values=row)
    return None


def clear_data():
    my_tree.delete(*my_tree.get_children())
    return None
root.mainloop()