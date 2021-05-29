
A -> promise date in the list of future
B -> request date in the list of future
C - > no promise date
D -> no request date
E 

if promise date not null
    promise date in the list of future
        put outstanding quantity in the cell of "Purchases Receipts + date"
    not in the list of future
        accumulate all this kind of value and put sum into the cell of "Purchases Receipts + past due"
else
    if request date is not null
        request date in the list of future
        put outstanding quantity in the cell of "All Open PO & BPO + date"
    not in the list of future
        accumulate all this kind of value and put sum into the cell of "All Open PO & BPO + past due"        
else
    all date null discard this record


    cell(row=Available Inventory, column=Past due) ---> W1 - demand 
    cell(row=Total Ending balance,column=past Due) ---> W1 + purchasde Receipts - demand

    row: Available Inventory + column :first day ---> cell(same_row, column -1).value -demand(current_day)  --> ending of week
    row: Total Ending balance + column :first day ---> cell(same_row, column -1).value  -demand(current_day) + purchase Receipt(current day)  --> ending of week
    
Genpak Data

    # "GP OH": 43008,       ->              past due
    # "WIP": null,                          GP WIP    
    # "On Order": 175000,   ->              GP OO        
    # "Target Date": "16-Jun-2021",         GP AVA DATE
    # "Ship Qy": "",                        FILM Release
    # "Del Date": ""                        Release date

Superpufft data
# "Current Stock/KG" ,       ->              past due
# "Next Available date & Quantity/KG",       GP WIP    


