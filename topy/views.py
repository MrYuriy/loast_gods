from asyncore import read
from operator import inv
from django.http import HttpResponse
from unicodedata import name
from urllib import response
from django.shortcuts import render
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
from django.http import JsonResponse, request, FileResponse
import os



def index(request):
    if "GET" == request.method:
        return render(request, 'topy/index.html',{})
    else :
        topy = core(request)
        file_path = "./topy.xlsx"
        response = HttpResponse(open(file_path, 'rb').read())
        print(response)
        response['Content-Type'] = 'mimetype/submimetype'
        response['Content-Disposition'] = 'attachment; filename=topy.xlsx'
        return response
def core(request):
    excel_file = request.FILES["excel_file"]
    wb = openpyxl.load_workbook(excel_file)
    
    
    
    transaction_sh = wb["transaction"]
    transactionarchiw_sh = wb["archiw"]
    inventory_sh = wb["inventory"]
    ean_sh = wb["sku"]
    topy_sh = wb["topy"]
    

    topy = get_sku_topy(topy_sh)
    transaction_archiw = read_transaction(transactionarchiw_sh)
    transaction = read_transaction(transaction_sh)
    inventory = get_inventory(inventory_sh)
    ean = get_ean(ean_sh)
    name = get_names(topy_sh)
    #print(name)
    #print(topy[0])
    #print(transaction[topy[1]],transaction_archiw[topy[1]],inventory[topy[1]],ean[topy[1]])
    #print(transaction[topy[1]],transaction_archiw[topy[1]])
    #topys = resoltb.active
    path = './topy.xlsx'
    try:
        os.remove(path)
        print("Deleted")
    except:
        None

    resoltb = Workbook()
    topys = resoltb.create_sheet("Topy",0)
    row = 1
    for top in topy:
        #top = 45887233
        # провірка на існування списку з адресами 
        if top in transaction and top in transaction_archiw:
            #print ('ok',transaction_archiw[top])
            adreses = unical_adres(transaction[top],transaction_archiw[top])
        elif top not in transaction and top in transaction_archiw:
            adreses = transaction_archiw[top]
        elif top in transaction and top not in transaction_archiw:
            adreses = transaction[top]
        else :
            adreses = []
        if inventory.get(top)!=None:
           aq = adresqty(adreses,inventory[top])
        else:
            aq = adreses
       #створення рядка для запису і одній клітинці всі адресів
        adrqtytowrite = ''
        name_ean = ""+name[top]+'\n'
        for a in aq:
            adrqtytowrite+=(a+'\n')
        #print(top,aq,ean[top])
        #print('\n')
        # додавання до адресів ean
        for e in ean[top]:
            name_ean+=(str(e)+'\n')
        #print(name_ean)
        #topys.cell('A1').style.alignment.wrap_text = True
        topys[f'C{row}'].alignment = Alignment(wrapText=True)
        topys[f'B{row}'].alignment = Alignment(wrapText=True)
        topys[f'A{row}']=top
        topys[f'B{row}']=name_ean
        #topys[f'B{row}'].alignment = Alignment(wrapText=True)
        topys[f'c{row}']=adrqtytowrite
        row+=1
    resoltb.save('topy.xlsx')
    print(type(resoltb))
    return (resoltb)    



def read_transaction(transaction_sh):
    sku_row = transaction_sh['B']
    from_location_row = transaction_sh['C']
    to_location_row = transaction_sh['D']
    transaction_dict = {}
    for s,fromloc,toloc in zip(sku_row,from_location_row,to_location_row):
        #print (s.value,'-',fromloc.value,'-',toloc.value)

        if s.value not in transaction_dict  :
            transaction_dict[s.value] = []
            if filter(fromloc.value):
                transaction_dict[s.value].append(fromloc.value)
            if filter(toloc.value):
                transaction_dict[s.value].append(toloc.value)
        else :
            if  fromloc.value not in transaction_dict[s.value] and filter(fromloc.value):
                transaction_dict[s.value].append(fromloc.value)
            if  toloc.value not in transaction_dict[s.value] and filter(toloc.value):
                transaction_dict[s.value].append(toloc.value)
    return (transaction_dict)
def get_sku_topy(topy_sh):
    sku_row = topy_sh['A']
    sku_list=[]
    for sku in sku_row:
        sku_list.append(sku.value) if type(sku.value)==int and sku.value not in sku_list else None
    
    return sku_list

def get_inventory(inventory_sh):
    sku_row = inventory_sh["F"]
    location_row = inventory_sh["G"]
    quty_row = inventory_sh["H"]
    # створюю не фільтрований  словник inventory_dict = {sku:[[adres,quty],[adres1,quty]]...sku_n:[[adres_n,quty],[adres_n+1,quty]]}
    inventory_dict_diorty = {}
    inventory_dict_clin = {}
    for s_row, l_row, q_row in zip(sku_row,location_row,quty_row):
        if s_row.value not in inventory_dict_diorty:
            if s_row.value != None:
                inventory_dict_diorty[s_row.value]=[[l_row.value,q_row.value]]
        else:
            if s_row.value != None:
                inventory_dict_diorty[s_row.value].append([l_row.value,q_row.value])
    # slq - sku + location + quanuty 
    # сумую співпвдіння по адресах ,1 адркс = загальна к-ть
    #print(inventory_dict_diorty)
    for slq in inventory_dict_diorty:
        betwin_dickt = {}
        for adres_qty in inventory_dict_diorty[slq]:
            #print(adres_qty[0])
            if adres_qty[0] not in betwin_dickt:
                betwin_dickt[adres_qty[0]]=adres_qty[1]
            else:
                # print('!!!!!!!!!!',adres_qty[0])
                # print("@@@@@@@@@@-",adres_qty[1])
                # print(adres_qty)
                betwin_dickt[adres_qty[0]]+=adres_qty[1]
        clean_list = []
        for adres in betwin_dickt:
            if filter(adres):
                clean_list.append([adres,betwin_dickt[adres]])
        inventory_dict_clin[slq]=clean_list
    # for i in inventory_dict_clin:
    #     print(i,'--',inventory_dict_clin[i])
    return (inventory_dict_clin)
   
def get_ean(ean_sh):
    sku_row = ean_sh["A"]
    ean_row = ean_sh["D"]
    ean_list = []
    suplier_dict = {}
    for s_row, e_row in zip(sku_row,ean_row):
        if e_row.value not in ean_list:
            ean_list.append(e_row.value)
            if s_row.value in suplier_dict :
                suplier_dict[s_row.value].append(e_row.value)
            else:
               suplier_dict[s_row.value]= [e_row.value]
    return(suplier_dict)
# обєднання адресів з врхіву і трансакшенів
def unical_adres(transaction , archiw):
    adres_list=archiw
    for adres in transaction:
        if adres not in archiw:
            adres_list.append(adres)
    return (adres_list)
# стврення списку  адресів і кілткості товару 
def adresqty(adreses, inventory):
    final_list = []
    not_add_list = []
    for aq in inventory:
        #print(aq)
        #print(aq[0]," - ",str(aq[1]))
        final_list.append(aq[0]+" - "+str(aq[1]))
        not_add_list.append(aq[0])
    for adres in adreses:
        if adres not in not_add_list:
            final_list.append(adres) 
    return (final_list)

def get_names(topy_sh):
    name_row = topy_sh['B']
    sku_row = topy_sh['A']
    name_dict = {}
    
    for sku, name in zip(sku_row,name_row):
        if type(sku.value)==int:
            name_dict[sku.value]=name.value
            #print(name.value)
    return name_dict

def filter(adres):
    #print(adres)
    adres = str(adres)
    adres = adres.replace('-','')
    if adres=='' or adres == None or len(adres)<=4:
        return False
    # куветки   K*******  K1204503  // вяти      W**** W0123 // карнизи   D**** D1330 
    if adres[0]=='K' or adres[0]=='W' or adres[0]=='D':
        if adres[1:].isdigit():
            return True
        else:
            return False

    elif (adres[0:1]=="WK" or adres[0:1]=="WS") and adres[2:].isgit():
        return True
            
    # регали    **R****L** 01R2104A10 01R1027A15 01R0517H1
    elif adres!="PARKING" and adres[2]=='R' and adres[7].isalpha():
        if adres[0:2].isdigit() and adres[3:7].isdigit():
            return True
        else:
            return False
    # маси      **M*** 01M242   01M525-20 01M316-03
    elif adres[2]=='M':
        if adres[3:].isdigit():
            return True
        else:
            return False
    elif adres[:8]=="SZYB-ZAM" :
        return True
    elif adres=='INRACK80':
        return True
    else:
        return False
        