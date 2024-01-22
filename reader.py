import pdfquery, os, re
import pandas as pd

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

columnas = ("CUIT","Cliente","Fecha", "Vto","Numero","Neto","Total","Tipo")

def main(path):

    info = dict(zip(columnas,([] for _ in range(len(columnas)))))
    for file in os.listdir(path):
        if file.endswith(".pdf"):
            numero = file.split()[0]
            if re.match(r"[0-9]*_001_[0-9]{5}_[0-9]*", numero):
                info = factura_a(file, info)
            elif re.match(r"[0-9]*_201_[0-9]{5}_[0-9]*", numero):
                info = fce_a(file,info)
            elif re.match(r"[0-9]*_006_[0-9]{5}_[0-9]*", numero):
                info = factura_b(file,info)
            elif re.match(r"[0-9]*_003_[0-9]{5}_[0-9]*", numero):
                info = nota_credito(file,info)

    df = pd.DataFrame.from_dict(info)
    df.to_excel(os.path.join(desktop,'facturas.xlsx'), sheet_name='sheet1', index=False)
                

def factura_a(file, info):
    pdf = pdfquery.PDFQuery(os.path.join(path, file))
    pdf.load()

    cuit = pdf.pq('LTTextBoxHorizontal:in_bbox("52.0, 642.69, 100.928, 650.69")')

    cliente = pdf.pq('LTTextBoxHorizontal:in_bbox("354.0, 642.69, 472.56, 650.69")')
    fecha = pdf.pq('LTTextBoxHorizontal:in_bbox("428.0, 730.36, 478.02, 740.36")')
    vto = pdf.pq('LTTextBoxHorizontal:in_bbox("495.0, 660.61, 545.02, 670.61")')
    numero = pdf.pq('LTTextBoxHorizontal:in_bbox("517.0, 745.86, 561.48, 755.86")')
    neto = float(pdf.pq('LTTextLineHorizontal:in_bbox("532.47, 279.7, 574.995, 288.7")')[0].text.strip().replace(",","."))
    total = neto + neto*0.21
    info["CUIT"].append(cuit[0].text)
    info["Cliente"].append(cliente[0].text)
    info["Fecha"].append(fecha[0].text)
    info["Vto"].append(vto[0].text)
    info["Numero"].append(int(numero[0].text))
    info["Neto"].append(neto)
    info["Total"].append(total)  
    info["Tipo"].append("Factura A")
    return info


def fce_a(file,info):

    pdf = pdfquery.PDFQuery(os.path.join(path, file))
    pdf.load()

    cuit = pdf.pq('LTTextBoxHorizontal:in_bbox("52.0, 645.69, 100.928, 653.69")')[0]
    cliente = pdf.pq('LTTextBoxHorizontal:in_bbox("354.0, 645.69, 465.44, 653.69")')
    fecha = pdf.pq('LTTextBoxHorizontal:in_bbox("428.0, 756.36, 478.02, 766.36")')
    vto = pdf.pq('LTTextBoxHorizontal:in_bbox("33.22, 687.61, 531.78, 697.61")')[0]
    vto = vto.text
    vto = vto[vto.find("Fecha de Vto. para el pago: "):vto.find("Per")].split()[-1]
    numero = pdf.pq('LTTextBoxHorizontal:in_bbox("517.0, 771.86, 561.48, 781.86")')
    neto = float(pdf.pq('LTTextLineHorizontal:in_bbox("530.47, 294.7, 577.999, 303.7")')[0].text.strip().replace(",","."))
    total = neto + neto*0.21
    
    info["CUIT"].append(cuit.text)
    info["Cliente"].append(cliente[0].text)
    info["Fecha"].append(fecha[0].text)
    info["Vto"].append(vto)
    info["Numero"].append(int(numero[0].text))
    info["Neto"].append(neto)
    info["Total"].append(total)
    info["Tipo"].append("FCE A")
    return info


def factura_b(file,info):
    pdf = pdfquery.PDFQuery(os.path.join(path, file))
    pdf.load()

    cuit = pdf.pq('LTTextBoxHorizontal:in_bbox("21.0, 635.52, 102.513, 644.52")')[0]

    cliente = pdf.pq('LTTextBoxHorizontal:in_bbox("353.0, 636.69, 557.256, 644.69")')
    fecha = pdf.pq('LTTextBoxHorizontal:in_bbox("428.0, 724.36, 478.02, 734.36")')
    vto = pdf.pq('LTTextBoxHorizontal:in_bbox("495.0, 654.61, 545.02, 664.61")')
    numero = pdf.pq('LTTextBoxHorizontal:in_bbox("517.0, 739.86, 561.48, 749.86")')
    neto = float(pdf.pq('LTTextBoxHorizontal:in_bbox("525.47, 214.52, 572.999, 223.52")')[0].text.strip().replace(",","."))
    total = neto + neto*0.21
    cuit = cuit.text.split()[1]
    info["CUIT"].append(cuit)   
    info["Cliente"].append(cliente[0].text)
    info["Fecha"].append(fecha[0].text)
    info["Vto"].append(vto[0].text)
    info["Numero"].append(int(numero[0].text))
    info["Neto"].append(neto)
    info["Total"].append(total)
    info["Tipo"].append("Factura B")
    return info


def nota_credito(file,info):
    pdf = pdfquery.PDFQuery(os.path.join(path, file))
    pdf.load()

    cuit = pdf.pq('LTTextBoxHorizontal:in_bbox("52.0, 642.69, 100.928, 650.69")')

    cliente = pdf.pq('LTTextBoxHorizontal:in_bbox("354.0, 642.69, 558.256, 650.69")')
    fecha = pdf.pq('LTTextBoxHorizontal:in_bbox("428.0, 730.36, 478.02, 740.36")')
    vto = pdf.pq('LTTextBoxHorizontal:in_bbox("495.0, 660.61, 545.02, 670.61")')
    numero = pdf.pq('LTTextBoxHorizontal:in_bbox("517.0, 745.86, 561.48, 755.86")')
    neto = float(pdf.pq('LTTextLineHorizontal:in_bbox("527.47, 279.7, 574.999, 288.7")')[0].text.strip().replace(",","."))
    total = neto + neto*0.21

    info["CUIT"].append(cuit[0].text)     
    info["Cliente"].append(cliente[0].text)
    info["Fecha"].append(fecha[0].text)
    info["Vto"].append(vto[0].text)
    info["Numero"].append(int(numero[0].text))
    info["Neto"].append(neto)
    info["Total"].append(total)
    info["Tipo"].append("NCE A")
    return info
    

if __name__ == "__main__":
    path = input("Path: ")
    main(fr"{path}")
