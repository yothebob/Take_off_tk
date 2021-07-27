from tkinter import *
import re
import openpyxl
from openpyxl import Workbook
from openpyxl.styles.borders import Border,Side
import math
from enum import Enum

"""
This was made to take in railing sections and give you an excel estimate based off the TR,BR,infill and etc you pick.
This also has the capability to create an estimate with total LF only also using the given TR,BR,infill. This is still in
trial and the total LF works based off statistics of previous estimates.



"""
#18-37 Enums for selecting TR,BR,Mount,infill
class TopRail:
    TOPRAIL = "Toprail"
    TR200 = "TR200"
    TR375 = "TR375"
    TR400 = "TR400"
    TR670 = "TR670"

class BottomRail:
    BOTTOMRAIL = "Bottomrail"
    BR200 = "BR200"
    BR100 = "BR100"

class Mount:
    MOUNT = "Mount"
    BP = "BP"
    FASCIA = "Fasica"
    HALFENS = "Halfen"
    COREHOLE = "Corehole"

class Infill:
    INFILL = "Infill"
    PICKET = "Picket"
    CABLE = "Cable"
    GLASS = "Glass"

def total_posts(_spacing):
    #return total number of posts and post types per spacing, only works with input_lengths
    sections = section.get('1.0',END)
    outside_corner = 0
    inside_corner = 0
    end_posts = 0
    line_posts = 0
    total_posts = 0
    total_sections = 0
    new_sect = sections.replace('\n',',')
    new_sect = list(new_sect.split(','))
    for line in new_sect:
        res = re.findall(r"\d?\d?[.]?\d?\d?",line)
        print(res)
        inside_corner += len(re.findall(r"[x]",line))
        outside_corner += len(re.findall(r"[*]",line))
        total_sections += 1
        for item in res:
            if item != "":
                total_posts += math.ceil(float(item)/ _spacing)
    total_posts += total_sections
    end_posts = total_sections *2
    corner_posts = outside_corner + inside_corner
    line_posts = total_posts - corner_posts - end_posts
    return total_posts,end_posts,corner_posts,line_posts


def return_spl200():
    #returns the amount of SPL200 parts needed
    spl200=0
    sections = section.get('1.0',END)
    new_sect= sections.replace('\n',',')
    new_sect=list(new_sect.split(','))
    for line in new_sect:
        res= re.findall(r"\d?\d?\.?\d\d?",line)
        for item in res:
            if float(item) > 20:
                spl200 += math.floor(float(item)/20)
    return spl200;


def total_tr(_tr_len=20):
    #totals and returns top railing, given a specific top rail length 
    sections = section.get('1.0',END)
    total_tr = 0
    total_scrap = []
    total_runs = []
    new_sect= sections.replace('\n',',')
    new_sect = list(new_sect.split(','))
    print('Totaling Top Rail...')
    for line in new_sect:
        res = re.findall(r"\d?\d?\.?\d\d?",line)
        for item in res:
            total_runs.append(float(item))
            total_scrap.append(0)

    total_runs = sorted(total_runs,reverse=True)
    print(total_runs)

    for run in range(len(total_runs)):
        print('RUN: ' + str(total_runs[run]))
        new_peice = True
        if total_runs[run] >= _tr_len:
            total_tr += math.ceil(float(total_runs[run])/_tr_len)
            total_scrap[run] = (math.ceil(float(total_runs[run])/_tr_len)*_tr_len) - total_runs[run]
        elif total_runs[run] < _tr_len:
            for scrap in range(len(total_scrap)):
                if total_runs[run] <= total_scrap[scrap]:
                    print('used scrap!')
                    print('scrap used: ' + str(total_scrap[scrap]))
                    print(total_scrap)
                    new_peice = False
                    total_scrap[scrap] -= total_runs[run]
                    break
            if new_peice == True:
                total_tr += 1
                total_scrap[run] = _tr_len - total_runs[run]

    return total_tr


def total_lf():
    # adds all sections together to return total LF
    lf=0
    sections = section.get('1.0',END)
    new_sect= sections.replace('\n',',')
    new_sect = list(new_sect.split(','))
    for line in new_sect:
        res = re.findall('\d?\d?\.?\d?\d',line)
        for item in res:
            lf += float(item)
    print("Total LF: " + str(lf))
    return lf


def make_xlsm(part_names,part_list):
    # make excel estimate given a list of parts and a list of part names
    wb=Workbook()
    color = openpyxl.styles.colors.Color(rgb='E5E7E6')
    my_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=color)
    thin_border = Border(bottom=Side(style='thin'))
    excel_name = take_off_name.get('1.0',END)
    excel_name = excel_name.replace('\n','')
    sheet = wb.active
    sheet['A1'] = excel_name
    sheet['B2'] = 'Part Names'
    sheet['C2'] = 'Quantity'
    sheet['D2'] = 'Field Measure'

    for num in range(len(part_list)):
        sheet.row_dimensions[num+3].height = 20
        sheet['B' + str(num+3)] = part_names[num]
        sheet['C'+ str(num+3)] = part_list[num]
        sheet['D' + str(num+3)].fill = my_fill
        sheet['B' + str(num+2)].border = thin_border
        sheet['C'+ str(num+2)].border = thin_border
        sheet['D'+ str(num+2)].border = thin_border
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['D'].width = 20

    wb.save(filename=excel_name + '.xlsx')
    print('Excel Sheet Made!!')


def total_parts_sections():
    #total parts using lf sections 
    part_names = ['Total LF','TR','BR','Pocket infill','Flat infill','SPE1', 'SPE2','PEU','p421','p422',
                  'fmpbs1','fmpbs2','Halfen','Corner Halfen','L bracket','pt-420','PVI','spacer 100',
                  'spacer 200','rcb1','rcb2','rcb screws','pc','gvs bot','gvs top','spl200','int90','int135',
                  'end plate','EP screw','SDS bag','NC/CW','lags',' Kwikset grout']

    selected_parts_dict = find_part_selections()

    job_lf = total_lf()
    tr = total_tr()
    br = tr

    #get post count and post spacing
    if selected_parts_dict[Infill.INFILL] == Infill.PICKET:#picket
        total_post,end_post,corner_post,line_post = total_posts(5)
    elif selected_parts_dict[Infill.INFILL] == Infill.GLASS:#glass
        total_post,end_post,corner_post,line_post = total_posts(4)
    elif selected_parts_dict[Infill.INFILL] == Infill.CABLE:#cable
        total_post,end_post,corner_post,line_post = total_posts(3)

    int90=0
    spl200= 0
    int90 = 0
    int135 = 0
    nccw = 0
    lags = 0
    pc = 0
    picket = 0
    pvi = 0
    spacer100 = 0
    spacer200=0
    sds_screw = 0
    rcb1 = 0
    rcb2 = 0
    rcb_screw = 0
    gvs_top = 0
    gvs_bot = 0
    pocket_infill = 0
    flat_infill = 0
    spe1 = 0
    spe2 = 0
    fmpbs1=0
    fmpbs2 = 0
    peu = 0
    grout = 0
    halfen = 0
    corner_halfen = 0
    l_bracket = 0
    end_plate = 0
    ep_screw = 0
    stair_bp = 0
    p421 = 0
    p422 = 0
    #----- mounting ------
    if selected_parts_dict[Mount.MOUNT] == Mount.BP: #baseplate
        nccw = total_post * 4
        lags = nccw
        if selected_parts_dict[TopRail.TOPRAIL] == TopRail.TR375:#tr375
            p421 = total_post
        else:
            p422 = total_post

    elif selected_parts_dict[Mount.MOUNT] == Mount.FASCIA: # fascia
        fmpbs1 = total_post - corner_post
        fmpbs2 = corner_post
        peu = math.ceil(total_post/4)
        nccw = total_post * 4
        lags = nccw

    elif selected_parts_dict[Mount.MOUNT] == Mount.HALFEN: # Halfen
        halfen = total_post - corner_post
        corner_halfen = corner_post
        l_bracket = total_post * 2
    else:
        peu = total_post/4
        grout = math.ceil(total_post / 12)
        
    #-----------infill type/ top rail type------------
    if selected_parts_dict[Infill.INFILL] == Infill.PICKET:#Picket
        picket = job_lf*3

        if selected_parts_dict[TopRail.TOPRAIL] == TopRail.TR375:#TR375
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
            pc = total_post
            sds_screw += ((total_post - end_post)*4) + (end_post *2)
        else:
            end_plate = end_post
            ep_screw = end_plate *2
            pocket_infill = br
            sds_screw += total_post * 4
            spl200 = return_spl200()
            int90 = corner_post

        if selected_parts_dict[BottomRail.BOTTOMRAIL] == BottomRail.BR200:#br200
            pvi = math.ceil((job_lf)/10)
            rcb2 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) * 2
            spacer100 = picket
            spacer200 = picket
        else:
            pvi = math.ceil((job_lf *2)/10)
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) * 2
            spacer100 = picket * 2
            sds_screw += ((total_post - end_post)*4) + (end_post *2)

    elif selected_parts_dict[Infill.INFILL] == Infill.GLASS:#glass
        gvs_top = math.ceil(job_lf/10)

        if selected_parts_dict[TopRail.TOPRAIL] == TopRail.TR375:#tr375
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
            pc = total_post
            sds_screw += ((total_post - end_post)*4) + (end_post *2)
        else:
            end_plate = end_post
            ep_screw = end_plate *2
            pocket_infill = tr
            sds_screw += ((total_post - end_post)*4) + (end_post *2)
            spl200 = return_spl200()
            int90 = corner_post

        if selected_parts_dict[BottomRail.BOTTOMRAIL] == BottomRail.BR200:#br200
            gvs_bot = gvs_top
            rcb2 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
        else:
            gvs_top = gvs_top * 2
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
            sds_screw += ((total_post - end_post)*4) + (end_post *2)

    elif selected_parts_dict[Infill.INFILL] == Infill.CABLE:#cable
        spacing = 3
        sds_screw += total_post * 4

        if selected_parts_dict[TopRail.TOPRAIL] == TopRail.TR375:#tr375
            rcb = total_post * 4
            pc = total_post
            spe1 = tr
        else:
            end_plate = end_post
            ep_screw = end_plate *2
            flat_infill = tr
            rcb = total_post * 2
            spl200 = return_spl200()
            int90 = corner_post

        rcb_screw = rcb * 2
        spe2 = br

    sds_screw = math.ceil(sds_screw /25)
    part_list = [job_lf,tr,br, pocket_infill,flat_infill,spe1, spe2,peu,p421,p422,fmpbs1,fmpbs2,
    halfen,corner_halfen,l_bracket,picket,pvi,spacer100,spacer200,rcb1,rcb2,rcb_screw,pc,gvs_bot,
    gvs_top,spl200,int90,int135,end_plate,ep_screw,sds_screw,nccw,lags,grout]
    print("Take Off Finished!!")
    make_xlsm(part_names,part_list)


def find_part_selections():
    #grab selected railing perameters and return an array with corrosponding numbers
    
    selected_mount =str_mount.get()
    selected_br = str_br.get()
    selected_tr = str_tr.get()
    selected_infill = str_infill.get()
    
    selected_parts_dict = {Mount.MOUNT: selected_mount,BottomRail.BOTTOMRAIL: selected_br, TopRail.TOPRAIL: selected_tr,Infill.INFILL : selected_infill}

    return selected_parts_dict



def total_parts_stats():
    """
    Uses some stats from C:\Users\Owner\Desktop\Estimating Tools\Python estimating tools\G drive Estimate scraping\cleaned data. 

    """

    
    part_names = ['Total LF','TR','BR','Pocket infill','Flat infill','SPE1', 'SPE2','PEU','p421',
                  'p422','fmpbs1','fmpbs2','Halfen','Corner Halfen','L bracket','pt-420','PVI',
                  'spacer 100','spacer 200','rcb1','rcb2','rcb screws','pc','gvs bot','gvs top',
                  'spl200','int90','int135','end plate','EP screw','SDS bag','NC/CW','lags',' Kwikset grout']

    job_lf = total_lf()
    #tr and bottom rail give 15% buffer
    tr = math.ceil((job_lf/20)*1.15)
    br = tr

    selected_parts_dict = find_part_selections()

    if selected_parts_dict[Infill.INFILL] == Infill.PICKET:#picket
        total_post = round(889.3+(-.335-889.3)/(1+(job_lf/1656.11)))
    elif selected_parts_dict[Infill.INFILL] == Infill.CABLE: #cable
        total_post = round(job_lf/2.4)
    elif selected_parts_dict[Infill.INFILL] == Infill.GLASS: # glass
        total_post = round(job_lf/3.1)

    end_post = round(121.3 + (-.3 - 121.3)/(1+(job_lf/344.14)))
    corner_post =round((.0167*job_lf)+1.5)
    line_post = total_post - end_post - corner_post

    int90=0
    spl200= 0
    int90 = 0
    int135 = 0
    nccw = 0
    lags = 0
    pc = 0
    picket = 0
    pvi = 0
    spacer100 = 0
    spacer200=0
    sds_screw = 0
    rcb1 = 0
    rcb2 = 0
    rcb_screw = 0
    gvs_top = 0
    gvs_bot = 0
    pocket_infill = 0
    flat_infill = 0
    spe1 = 0
    spe2 = 0
    fmpbs1=0
    fmpbs2 = 0
    peu = 0
    grout = 0
    halfen = 0
    corner_halfen = 0
    l_bracket = 0
    end_plate = 0
    ep_screw = 0
    stair_bp = 0
    p421 = 0
    p422 = 0

    #------------------------Mounting--------------------------------
    if selected_parts_dict[Mount.MOUNT] == Mount.BP: #baseplate
        nccw = total_post * 4
        lags = nccw
        if selected_parts_dict[TopRail.TOPRAIL] == TopRail.TR375:#tr375
            p421 = total_post
        else:
            p422 = total_post

    elif selected_parts_dict[Mount.MOUNT] == Mount.FASCIA: # fascia
        fmpbs1 = total_post - corner_post
        fmpbs2 = corner_post
        peu = math.ceil(total_post/4)
        nccw = total_post * 4
        lags = nccw

    elif selected_parts_dict[Mount.MOUNT] == Mount.COREHOLE: # corehole
        peu = total_post/4
        grout = math.ceil(total_post / 12)
    else:
        halfen = total_post - corner_post
        corner_halfen = corner_post
        l_bracket = total_post * 2
#--------------------------infill / TR type-------------------------------
    if selected_parts_dict[Infill.INFILL] == Infill.PICKET:#Picket
        picket = job_lf*3

        if selected_parts_dict[TopRail.TR375] == TopRail.TR375:#TR375
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
            pc = total_post
            sds_screw += ((total_post - end_post)*4) + (end_post *2)
        else:
            end_plate = end_post
            ep_screw = end_plate *2
            pocket_infill = br
            sds_screw += total_post * 4
            spl200 = 0#return_spl200()
            int90 = corner_post

        if selected_parts_dict[BottomRail.BOTTOMRAIL] == BottomRail.BR200:#br200
            pvi = math.ceil((job_lf)/10)
            rcb2 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) * 2
            spacer100 = picket
            spacer200 = picket
        else:
            pvi = math.ceil((job_lf *2)/10)
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) * 2
            spacer100 = picket * 2
            sds_screw += ((total_post - end_post)*4) + (end_post *2)

    elif selected_parts_dict[Infill.INFILL] == Infill.GLASS:#glass
        gvs_top = math.ceil(job_lf/10)

        if selected_parts_dict[TopRail.TOPRAIL] == TopRail.TR375:#tr375
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
            pc = total_post
            sds_screw += ((total_post - end_post)*4) + (end_post *2)
        else:
            end_plate = end_post
            ep_screw = end_plate *2
            pocket_infill = tr
            sds_screw += ((total_post - end_post)*4) + (end_post *2)
            spl200 = 0#return_spl200()
            int90 = corner_post

        if selected_parts_dict[BottomRail.BOTTOMRAIL] == BottomRail.BR200:#br200
            gvs_bot = gvs_top
            rcb2 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
        else:
            gvs_top = gvs_top * 2
            rcb1 += ((total_post - end_post) * 2) + (end_post)
            rcb_screw += (((total_post - end_post) * 2) + (end_post)) *2
            sds_screw += ((total_post - end_post)*4) + (end_post *2)


    elif selected_parts_dict[Infill.INFILL] == Infill.CABLE:#cable
        spacing = 3
        sds_screw += total_post * 4

        if selected_parts_dict[TopRail.TOPRAIL] == TopRail.TR375:#tr375
            rcb = total_post * 4
            pc = total_post
            spe1 = tr
        else:
            end_plate = end_post
            ep_screw = end_plate *2
            flat_infill = tr
            rcb = total_post * 2
            spl200 = 0#return_spl200()
            int90 = corner_post

        rcb_screw = rcb * 2
        spe2 = br

    sds_screw = math.ceil(sds_screw /25)
    part_list = [job_lf,tr,br, pocket_infill,flat_infill,spe1, spe2,peu,p421,p422,fmpbs1,fmpbs2,
    halfen,corner_halfen,l_bracket,picket,pvi,spacer100,spacer200,rcb1,rcb2,rcb_screw,pc,gvs_bot,
    gvs_top,spl200,int90,int135,end_plate,ep_screw,sds_screw,nccw,lags,grout]
    print("Take Off Finished!!")
    make_xlsm(part_names,part_list)



#creating gui for running 
window=Tk()
window.title('Take Off Tool 0.2')

str_tr= StringVar(window)
str_br= StringVar(window)
str_mount= StringVar(window)
str_infill= StringVar(window)

str_tr.set(TopRail.TR200)
str_br.set(BottomRail.BR200)
str_mount.set(Mount.BP)
str_infill.set(Infill.PICKET)

section = Text(width=80,height=25)
section.grid(row=1,column=0)
sect_label= Label(text='Sections: (10x10) for inside corner use *').grid(row=0,column=0)
tr_box = OptionMenu(window,str_tr,'TR200','TR375','TR400','TR670').grid(row=2,column=0)
br_box = OptionMenu(window,str_br,'BR200','BR100').grid(row=3,column=0)
mount_box= OptionMenu(window,str_mount,'BP','Foam Blockout','Halfen','Fascia').grid(row=4,column=0)
infill_box= OptionMenu(window,str_infill,'Picket','Glass','Cable').grid(row=5,column=0)
take_off_label = Label(text='Job Name:').grid(row=6, column=0)
take_off_name = Text(width=30,height=1)
take_off_name.grid(row=7,column=0)
enter_section_button = Button(text='Run (Sections ONLY)',command=total_parts_sections).grid(row=8,column=0)
enter_stat_button = Button(text="Run (LF ONLY)",command=total_parts_stats).grid(row=9,column=0)
window.mainloop()

