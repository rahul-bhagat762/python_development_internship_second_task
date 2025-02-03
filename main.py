import tkinter
from tkinter import ttk #this is used for user interface
from docxtpl import DocxTemplate #for handling microsoft doc template
import datetime #for datatime

def clear_item():
    qty_entry.delete(0,tkinter.END)
    qty_entry.insert(0,"1")
    desc_entry.delete(0,tkinter.END)
    price_entry.delete(0,tkinter.END)
    price_entry.insert(0,"0.0")

invoice_list = []
def add_item():
    qty = int(qty_entry.get())
    desc = desc_entry.get()
    price = float(price_entry.get())
    line_total = qty*price
    invoice_item = [qty,desc,price,line_total]
    tree.insert('',0,values =invoice_item)
    clear_item()

    invoice_list.append(invoice_item)

def new_invoice():
    first_name_entry.delete(0,tkinter.END)
    last_name_entry.delete(0,tkinter.END)
    phone_entry.delete(0,tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())

    invoice_list.clear()


def generate_invoice():
    doc  = DocxTemplate('invoice_template_update.docx') ##first Import docx template, now its time to load tempalate and this line is used to load tempate
    name = first_name_entry.get()+last_name_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax =0.1
    total = subtotal*(1-salestax)
    company = 'ABC PVT LTD.'

    doc.render({'name':name,
                "Company" : company,
                "phone":phone,
                "invoice_list":invoice_list,
                "subtotal":subtotal,
                "salestax":str(salestax*100)+"%",
                "total":total})   ## To actually make any chnages you need to render first

    doc_name = "new_invoice"+ name +datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name) ##After apply various operations and now its time to save final product


# this portion is for user interface
## This is a whole screen 
window = tkinter.Tk() ##Basically i create here my root widget which makes window of our project
window.title("Invoice Generator Form") ##This line is used to print title of my window 

##This below line is used to make container(frame)
# This frame is small box of a bigger window
frame = tkinter.Frame(window)
frame.pack() ##after create we have to pack in this frame

##This is a small box for entry , this is a part of a frame
first_name_label = tkinter.Label(frame,text="First Name") ##This text attribute is used to define what we want look in this box
first_name_label.grid(row=0,column=0) ##his is used for where you want to fix this label
last_name_label = tkinter.Label(frame,text="Last Name") ##This text attribute is used to define what we want look in this box
last_name_label.grid(row=0,column=1)

first_name_entry = tkinter.Entry(frame)
last_name_entry  = tkinter.Entry(frame)
first_name_entry.grid(row=1,column=0)
last_name_entry.grid(row=1,column=1) 

phone_label = tkinter.Label(frame,text='Phone')
phone_label.grid(row=0,column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1,column=2)

qty_label = tkinter.Label(frame,text='Qty')
qty_label.grid(row=2,column=0)
qty_entry = tkinter.Spinbox(frame,from_=1,to=100) ##this line is used for spinbox 
qty_entry.grid(row=3,column=0)

desc_label = tkinter.Label(frame,text='Discription')
desc_label.grid(row=2,column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3,column=1)

price_label = tkinter.Label(frame,text='Unit Price')
price_label.grid(row=2,column=2)
price_entry = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_entry.grid(row=3,column=2)

add_item_button = tkinter.Button(frame,text="Add Item",command= add_item)
add_item_button.grid(row=4,column=2, pady=5)

columns = ('qty','desc','price','total')
tree = ttk.Treeview(frame,columns=columns,show='headings')
tree.heading('qty',text='Qty')
tree.heading('desc',text='Description')
tree.heading('price',text='Unit Price')
tree.heading('total',text='Total')

tree.grid(row = 5,column= 0,columnspan=3,padx=20,pady=10)

save_invoice_button = tkinter.Button(frame,text="Generate Invoice",command=generate_invoice)
save_invoice_button.grid(row=6,column=0,columnspan=3,sticky="news",padx = 20,pady=5)
new_invoice_button = tkinter.Button(frame,text="New Invoice",command=new_invoice)
new_invoice_button.grid(row=7,column=0,columnspan=3,sticky="news",padx = 20,pady=5)


window.mainloop() ##This is a infite loop , this loop run infinite time when this app will close then this loop is terminate


##Thank You Pls Like and follow me