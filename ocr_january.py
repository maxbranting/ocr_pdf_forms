import PyPDF2
import pypdfium2 as pdfium
import pytesseract
from PIL import Image
from numpy import mean, array
import pandas as pd
from skimage.filters import threshold_sauvola
import datetime as dt
import string
import os
from pathlib import Path
import re
import traceback
from variables import output_template, version_tag, coordinates


# pytesseract.pytesseract.tesseract_cmd = r'D:\ProgramData\Tesseract_OCR\tesseract'


pytesseract.pytesseract.tesseract_cmd = r'C:\Users\UI954923\AppData\Local\Tesseract-OCR\tesseract'


pth = r'C:\Users\UI954923\Desktop\vsc_repositories\frankenstein_repo\pdfstore' # folder z plikami pdf do procesowania
destpth = r'C:\Users\UI954923\Desktop\vsc_repositories\frankenstein_repo\pdfdone' # folder z przeprocesowanymi pdfami
outpth = r'C:\Users\UI954923\Desktop\vsc_repositories\frankenstein_repo\ocrout' # folder z outputem csv
err_mes = r'C:\Users\UI954923\Desktop\vsc_repositories\frankenstein_repo\ERROR\Message' # folder z zapisem bledow
err_move = r'C:\Users\UI954923\Desktop\vsc_repositories\frankenstein_repo\ERROR\Files' 


paths = []
for file in os.listdir(pth):
    paths.append(file) # lista pdfów w ściece

# init słownika output
output_main = {}


def clean_out(output):
    df = pd.DataFrame(output).T.reset_index() # Transpose otputu
    df1 = df.copy()
    # remove /n
    df1 = df.replace(r'\n','', regex=True) 
    # signature date
    df1['SignatureDate'] = pd.to_datetime(df1['SignatureDate'].str.extract("(\d+)", expand=False), format='%Y%m%d%H%M%S') # podpisy ele mają w metadanych timestampy
    # PESEL
    df1['PESEL'] = df1['PESEL'].str.translate(str.maketrans('', '', string.whitespace)) # to i kolejne ma whitespacey których nie da się pozbyć innymi metodami
    # NR Konta Umowy
    df1['nrKU'] = df1['nrKU'].str.translate(str.maketrans('', '', string.whitespace))
    # KARTA DUŻEJ RODZINY
    df1['nrKDR'] = df1['nrKDR'].str.translate(str.maketrans('', '', string.whitespace))
    # NR PPE
    df1['nrPPE'] = df1['nrPPE'].str.translate(str.maketrans('', '', string.whitespace))
    df1['nrPPE'] = df1['nrPPE'].str.replace('O', '0') # tesseract czasami myli zero z O
    # column order
    colord = ['GUID', 'filename', 'Signature1', 'SignatureDate', 'imie', 'nazwisko', 'PESEL', 'nrKU', 'adreskoresp_adresppe', 'adresPPE', 'nrPPE', 'cbox_rolnik', 'cbox_rodzina', 'cbox_niepeln', 'cbox_budowa', 'cbox_dzialka', 'gra', 'grb', 'nrKDR', 'l_dzialekPPE', 'dzien1', 'miesiac1', 'rok1', 'data_nab_uprawn', 'do_weryfikacji']
    df1 = df1[colord]

    return df1

class MoveFile(Exception):
    '''Helper exception'''
    pass

def pdpage_to_image(image, page_n):
    """Helper function for converting pdf page into an image and thresholding it"""
    page = image.get_page(page_n)
    pil_image = page.render_topil(
            scale=4.5,
            rotation=0,
            greyscale=True,
            optimise_mode=pdfium.OptimiseMode.NONE,
    )

    nimg = array(pil_image) # convert to np array for skimage
    thresh = threshold_sauvola(nimg) # thresholding with skimage sauvola

    fimg = nimg > thresh # binarization of image
    pfimg = Image.fromarray(fimg).convert('RGB') # back to pil format 
    return pfimg



for pdf in paths:
    
    print(pdf)
    output_main = {}
    output_inside = output_template.copy()

    try:
        with open(os.path.join(pth, pdf), 'rb') as file:
            document = PyPDF2.PdfFileReader(file, strict=False)
            fields = document.get_fields()
            
            tickkey = ['rolnik', 'rodzina', 'niepeln', 'dzialka']
            datekey = ['dzien1', 'miesiac1', 'rok1', 'kod']
            outkey = [x for x in list(output_inside.keys())[2:] if x not in tickkey and datekey] # lista pól które nie mają tickmarków ani dat rozdzielonych na 

            splt = re.split('_|\.', pdf) # rodzielenie nazwy pliku na podst _ i kropki
            guid = splt[0] # pierwsza część to guid
            filename = splt[1] # druga nazwa pliku

            output_inside['filename'] = filename 
            output_inside['GUID'] = guid     # przypisanie zmiennych do odp kolumn
            
            try:  # error dealing z pdfami które są zeskanowane 
                pdf_keys = fields.keys()

            except AttributeError: # .keys wyrzuca AttributeError na skanach nie podpisanych elektronicznie
                output_inside['Signature1'] = False

                for x in outkey[2:]:
                    output_inside[x] = None
                
                output_inside['do_weryfikacji'] = True
                output_main[pdf] = output_inside # filling main dictionary 

                mout = clean_out(output_main)  # czyszczenie funkcją pomocniczą
                mout.to_csv(outpth +'/'+ pdf[:-4]+'.csv', sep=';', index=False) # zapis pojedyńczych plików
                
                raise MoveFile
            
            for key in pdf_keys:
                sign = ''
                if 'signature' in key.lower():
                    sign = key

            if sign not in pdf_keys: # sprawdzanie czy podpis obecny, wyciąganie jego daty
                output_inside['Signature1'] = False

                for x in outkey[2:]:
                    output_inside[x] = None
                
                output_main[pdf] = output_inside
                mout = clean_out(output_main)  # czyszczenie funkcją pomocniczą
                mout.to_csv(outpth +'/'+ pdf[:-4]+'.csv', sep=';', index=False) 
                
                raise MoveFile

            else:
                output_inside['Signature1'] = True
                output_inside['SignatureDate'] = fields[sign]['/V']['/M']
            
                with open(os.path.join(pth, pdf), 'rb') as file:
                    pdfim = pdfium.PdfDocument(file)
                    #n_pages = len(pdf)
                    n_pages = 2

                    try:
                        vimg = pdpage_to_image(pdfim, 0)
                        version_full = str(pytesseract.image_to_string(vimg.crop(version_tag[0]), config = '--psm 6 --oem 3'))  # selecting subset coordinates for correct version of form
                        version_data = version_full.split(' ')[-1].split('nr',1)[0].strip('\n')
                        version = version_data[4:8]
                        assert version=='1115' or version=='1104' or version=='1019' or version=='1027' or version=='0102'

                    except:
                        for x in outkey[2:]:
                            output_inside[x] = None

                        output_inside['do_weryfikacji'] = True
                        output_main[pdf] = output_inside
                        mout = clean_out(output_main)  # czyszczenie funkcją pomocniczą
                        mout.to_csv(outpth +'/'+ pdf[:-4]+'.csv', sep=';', index=False) 

                        raise MoveFile


                    for page_number in range(n_pages): # conversion of pdf page to an image
                            pfimg = pdpage_to_image(pdfim, page_number)

                            for index, coord in coordinates[version][page_number].items():
                                    imgt = pfimg.crop(coord)  # cropping a region of page an check with below conditions
                                    
                                    if index in ['adreskoresp_adresppe', 'cbox_rolnik', 'cbox_rodzina', 'cbox_niepeln', 'cbox_budowa', 'cbox_dzialka']: # loop for checking if tick is present by taking avg col of space
                                        if mean(array(imgt)) <=240:
                                            output_inside[index] = True
                                        else:
                                            output_inside[index] = False 

                                    elif index in ['PESEL', 'nrKU', 'nrPPE']:
                                        output_inside[index] = str(pytesseract.image_to_string(imgt, config = '--psm 6 --oem 3 -c tessedit_char_blacklist="Oo"'))
                                    elif index in ['data_nab_uprawn', 'l_dzialekPPE']:
                                        output_inside[index] = str(pytesseract.image_to_string(imgt, config = '--psm 6 --oem 3'))
                                    else:
                                        output_inside[index] = str(pytesseract.image_to_string(imgt, lang='pol', config = '--oem 3')) # going through non tick fields and extracting text   ##config = '--psm 13'

            if len("".join(re.findall("[a-zA-Z]+", output_inside['PESEL']))) > 0: # ta sekcja sprawdza czy jest bełkot w danych liczbowych
                output_inside['do_weryfikacji'] = True
            elif len("".join(re.findall("[a-zA-Z]+", output_inside['nrKU']))) > 0:
                output_inside['do_weryfikacji'] = True
            elif len("".join(re.findall("[a-zA-Z]+", output_inside['nrPPE']))) > 2: # ten numer może zawierać dwie litery
                output_inside['do_weryfikacji'] = True
            else:
                output_inside['do_weryfikacji'] = False


            output_main[pdf] = output_inside # filling main dictionary 

        mout = clean_out(output_main)  # czyszczenie funkcją pomocniczą
        mout.to_csv(outpth +'/'+ pdf[:-4]+'.csv', sep=';', index=False) # zapis pojedyńczych plików

        Path(pth+'/'+pdf).rename(destpth+'/'+pdf) # przenoszenie pliku pdf do drugiego folderu

    except MoveFile:
        Path(pth+'/'+pdf).rename(destpth+'/'+pdf)

    except Exception as e:
        mess = traceback.format_exc()
        with open(err_mes + '/' + pdf + '.txt', 'w') as text_file:
            text_file.write(str(mess))
        Path(pth+'/'+pdf).rename(err_move + '/' +pdf)