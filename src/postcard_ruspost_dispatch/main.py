#!/usr/bin/env python3
# -*- coding: utf-8 -*-

#pip3 install docxtpl
#pip3 install Pillow
#pip3 install python-barcode


import re
import os
import io

from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document as Document_compose
from docx.shared import Mm

from PIL import Image, ImageDraw, ImageFont

import barcode
from barcode.writer import ImageWriter

import post_rf_barcodes_lib
import sender_config

import subprocess

def create_list(file_name):
  results = []
  expected_name_index = 1
  with open(file_name, 'r') as f:
    for item in list(f):
      if not (item and item.strip()) :
        print ("Error in line " + str(len(results)-1) + " : line is empty." ) ;
        continue
      
      item = re.sub('[!@#$\n]', '', item)
      parse_item = item.split("\t")
      # print (parse_item)
      p_type = define_name_association_form (parse_item[expected_name_index], parse_item[0].lower())
      
      parse_name = parse_item[expected_name_index].split()
      # print (parse_name)
      
      # set gender
      if re.match(".*[вч]на", parse_name[2]): p_male = "girl"
      if re.match(".*[вь]ич", parse_name[2]): p_male = "boy"
      if re.match(".* оглы",  parse_item[0]): p_male = "boy"
      
      # extract postcode
      p_pocstcode = ""
      p_address   = ""
      index = re.findall("[0-9]{6}", parse_item[expected_name_index+1])
      if len(index) == 0:
        print ("Error in line " + str(len(results)+1) + " : not find index." ) ;
      else:
        p_pocstcode = int(index[0].strip())
        p_address   = re.sub("[0-9]{6}, ", "", parse_item[expected_name_index+1])
        p_address   = re.sub("\(.*\)", "", p_address)
      
      # print (str(p_pocstcode) + "--" + str(p_address))
      p_track = []
      for x in range(int(parse_item[expected_name_index+5])):
        if len(parse_item)>=expected_name_index+8+x:
          p_track.append(parse_item[expected_name_index+8+x].strip())
      if len(p_track)==0 : p_track.append(0)
      
      # calculate waight
      p_waight = round(int(parse_item[expected_name_index+4].strip())/len(p_track))
      # calculate cost
      
      p_cost = float(parse_item[expected_name_index+6].replace(',', '.').strip())/len(p_track)
      
      for px_track in p_track:
        results.append({
          'destination_surname'     : parse_name[0].strip(),
          'destination_name'        : parse_name[1].strip(),
          'destination_father'      : parse_name[2].strip(),
          'destination_bdate'       : parse_name[len(parse_name)-2].strip(),
          'destination_association' : p_type,
          'destination_male'        : p_male,
          'destination_code'        : p_pocstcode,
          'destination_addr'        : p_address,
          'destination_waight'      : p_waight,
          'destination_cost'        : p_cost,
          'destination_track'       : px_track,
          'destination_cnt'         : len(p_track),
          })
        
  return results

def define_name_association_form(destination_name, destination_prefix):
  # set name association form
  p_type = 0
  if not (destination_prefix in ["ты", "вы"]) :
    print ("Error in name \"" + destination_name + "\" : not set prefix." ) ;
    return 0 # вы - отчество

  destination_year = 0 
  year = re.findall("[0-9]{4}", destination_name)
  if len(year) == 0:
    print ("Error in name \"" + destination_name + "\" : not parse year." ) ;
  else:
    destination_year = int(year[0].strip())
  
  if destination_prefix == "вы" : p_type = 0 # вы - отчество
  if destination_prefix == "ты" : p_type = 1 # ты - без отчества
  if (destination_year > 1984) & (p_type == 0) : p_type = 2 # вы - без отчества
  return p_type

def replace_by_template(destination):
  p_full_name = destination['destination_surname'] + " " + \
    destination['destination_name']    + " " + \
    destination['destination_father']  + " " + \
    destination['destination_bdate']

  p_denotation = destination['destination_name']
  if destination['destination_association'] == 0 : p_denotation += " " + destination['destination_father']
  p_surname = destination['destination_surname']
  
  if destination['destination_association'] == 1 :
    doc = DocxTemplate("letter_template_relative.docx")
  else :
    doc = DocxTemplate("letter_template.docx")
  new_file_name = "let_" + p_surname + ".docx"
  
  p_one = "один"
  if destination['destination_male'] == 'girl' : p_one = "одна"

  context = {'prisoner_full_data' : p_full_name, 'prisoner_name' : p_denotation, 'prisoner_one' : p_one }
  doc.render(context)
  doc.save(new_file_name)
  return new_file_name

def check_duplicates(p_list):
  print ("check dupl")
  for destination in p_list:
    match_list = [ i for i in p_list if destination['destination_surname'] == i['destination_surname'] ]
    # print (match_list)
    if len(match_list)>destination['destination_cnt'] :
      print ("list contains duplicates of \"" + destination['destination_surname'] + "\"")

def create_inserted_document(destination_list, dst_fold_name):
  #filename_master is name of the file you want to merge the docx file into
  master = Document_compose()

  master.sections[0].page_height = Mm(297)
  master.sections[0].page_width = Mm(210)
  master.sections[0].left_margin = Mm(15.4)
  master.sections[0].right_margin = Mm(15.4)
  master.sections[0].top_margin = Mm(15.4)
  master.sections[0].bottom_margin = Mm(15.4)
  master.sections[0].header_distance = Mm(12.7)
  master.sections[0].footer_distance = Mm(12.7)
  composer = Composer(master)

  for destination in destination_list: 
    f_name = replace_by_template (destination)

    #filename_second_docx is the name of the second docx file
    doc2 = Document_compose(f_name)
    #append the doc2 into the master using composer.append function
    composer.append(doc2)
    os.remove(f_name)

  #Save the combined docx with a name
  composer.save( dst_fold_name + "/covering_letter_full.docx")

def split_line_by_length(init_line, max_length=28, skew=6):
  word_array = init_line.split()
  part_array = []
  
  i=0
  for word in word_array:
    i+=1
    if word in ["г.", "г", "ул", "ул.", "д.", "д", "ФКУ", "по", "пгт.", "пгт", "мкр.", "мкр"]:
      word_array[i] =' '.join([ word, word_array[i] ])
      word_array.remove(word)
  
  cur_len = 0
  cur_line = ""
  
  for part in word_array:
    if cur_len + len(part) <= max_length:
      cur_len += len(part) + 1 
      cur_line += str(part) + " "
    else:
      cur_len = len(part) + 1 
      cur_line += "\n" + ''.rjust(skew, ' ') + str(part) + " "
  
  return cur_line

def create_envelope(destination, font_fold_name, dst_fold_name, i):
  filename_base = str(dst_fold_name) + str(i).zfill(2) + "_" + str(destination['destination_surname']) + "_" + str(destination['destination_name']) + "_" + str(destination['destination_track'])
  
  from_form = 'От кого: ' + sender_config.FROM_CONF
  
  to_form = "Кому: " + split_line_by_length (str(destination['destination_surname']) + " " \
                       + str(destination['destination_name']) + " " \
                       + str(destination['destination_father']) + " " \
                       + str(destination['destination_bdate'])) + \
      "\n\nКуда: " + split_line_by_length (str(destination['destination_addr']))
  
  mark_form = str(destination['destination_waight']) + " гр.\n" +  str(destination['destination_cost']) + " р."
  
  # c6 : 114 × 162
  img = Image.new('RGB', (1620, 1140), color = (255, 255, 255))
  
  fnt = ImageFont.truetype(font_fold_name + '/AnonymousPro-Regular.ttf', 50)
  add_fnt = ImageFont.truetype(font_fold_name + '/AnonymousPro-Regular.ttf', 30)
  ind_fnt = ImageFont.truetype(font_fold_name + '/PostIndex.ttf', 120)
  
  d = ImageDraw.Draw(img)
  d.text((40,50), from_form, font=fnt, fill=(0, 0, 0))
  d.text((600,600), to_form, font=fnt, fill=(0, 0, 0))
  d.text((1400,80), mark_form, font=add_fnt, fill=(100, 100, 100))
  
  # print barcode to an actual file:
  if destination['destination_track'] != '' and destination['destination_track'].isnumeric():
    with open(filename_base + 'track.png', 'wb') as f:
      options = {
        'dpi': 500,
        'module_height': 8,
        'quiet_zone': 1,
        'text_distance': 4,
        'font_size': 8
      }
      ITF = barcode.get_barcode_class('itf')
      ITF(destination['destination_track'], writer=ImageWriter()).write(f, options = options)
      # EAN14(destination['destination_track'], writer=ImageWriter()).write(f)
    barkode = Image.open(filename_base + 'track.png')#.resize((600, 200))
    os.remove(filename_base + 'track.png')
    img.paste(barkode, (40, 600))
  
  # draw postcode
  d.text((50,900), '\U00000024' + str(destination['destination_code']), font=ind_fnt, fill=(0, 0, 0))
  
  img.save(filename_base + "_envelope.png")
  return

def debug_print_list(p_list):
  for destination in p_list:
    print('%10s : %4s : %1s : %3s : %2s : %6s : %14s '\
      % ( str(destination['destination_surname']),\
        str(destination['destination_male']),\
        str(destination['destination_association']),\
        str(destination['destination_waight']),\
        str(destination['destination_cost']),\
#       + str(destination['destination_association']) + " : " \
       # + str(destination['destination_addr']) + " : " \
        str(destination['destination_code']),\
        str(destination['destination_track'])))
        
  return


def check_tracknumber_link(destination):

  if destination['destination_track'] != '' :
    for barcode in [destination['destination_track']]:
      extracted_name = post_rf_barcodes_lib.expract_name_by_barcode( barcode )
      if len (post_rf_barcodes_lib.error_barcodes) > 0 :
        print ("ERROR! Barcode " + str(destination['destination_track']) +  " doesn't exist - " + str(destination['destination_surname']))
        return 0
      print(extracted_name)
     
      expected_name = []
      expected_name.append ((str(destination['destination_surname']) + " " + str(destination['destination_name']) + " " + str(destination['destination_father'])).upper())
      expected_name.append ((str(destination['destination_surname']) + " " + str(destination['destination_name'])[0] + ". " + str(destination['destination_father'])[0]).upper() + ".")
      expected_name.append ((str(destination['destination_surname']) + " " + str(destination['destination_name']) + " " + str(destination['destination_father']) + str(destination['destination_bdate'])).upper())
        
      if not (extracted_name.upper() in expected_name) :
        print ("ERROR! Extracted " + extracted_name + " for " + str(destination['destination_track']) + " not equal to expected " + expected_name[0])
        return 0
  return 1

def create_barcode_document (destination_list, font_fold_name, dst_fold_name ):
  filename_base = dst_fold_name + "/list"
  
  if not os.path.exists(dst_fold_name):
    print("Create dst_fold_name folder: \"" + dst_fold_name + "\"")
    os.makedirs(dst_fold_name)
  for file_name in os.listdir(dst_fold_name):
    if os.path.isfile(os.path.join(dst_fold_name, file_name)):
      os.remove(os.path.join(dst_fold_name, file_name))
  
  x_ind=0
  y_ind=0
  f_ind=0
  x_border = 80
  y_len = 330
        # 3x9 : 7*3 
  img = Image.new('RGB', (3*700, 9*y_len), color = (255, 255, 255))
  fnt = ImageFont.truetype(font_fold_name + "/AnonymousPro-Regular.ttf", 20)
  d = ImageDraw.Draw(img)
        
  for destination in destination_list:
    # print barcode to an actual file:
    if destination['destination_track'] != '' and destination['destination_track'].isnumeric() :
      with open(filename_base + 'track.png', 'wb') as f:
        options = {
          'dpi': 500,
          'module_height': 8,
          'quiet_zone': 1,
          'text_distance': 2,
          'font_size': 6
        }
        ITF = barcode.get_barcode_class('itf')
        ITF(destination['destination_track'], writer=ImageWriter()).write(f, options = options)
      barkode = Image.open(filename_base + 'track.png')
      os.remove(filename_base + 'track.png')
      img.paste(barkode, (x_border+x_ind*700, y_ind*y_len))
    
      d.text((x_border+x_ind*700,y_ind*y_len+230), str(destination['destination_surname']) + " " + str(destination['destination_name']), font=fnt, fill=(0, 0, 0))
      if x_ind < 2: x_ind +=1
      else: 
        x_ind = 0
        y_ind +=1
      if y_ind == 9 :
        img.save(filename_base + "_" + str(f_ind) + ".png")
        img = Image.new('RGB', (3*700, 9*y_len), color = (255, 255, 255))
        d = ImageDraw.Draw(img)
        y_ind = 0
        f_ind +=1
  img.save(filename_base + "_" + str(f_ind) + ".png")
  return


def main():
  d=os.path.dirname
  _TOP_FOLD = d(d(d(os.path.abspath(__file__))))
  _SRC_LIST = os.path.join(_TOP_FOLD, "prisoner_list.txt" )
  _OUT_FOLD = os.path.join(_TOP_FOLD, "out" )
  
  print(_TOP_FOLD) 
  print(_SRC_LIST) 
  if not os.path.exists(_SRC_LIST):
    print("ERROR: file \"" + _SRC_LIST + "\" not found")
    return
    
  if not os.path.exists(_OUT_FOLD):
    print("Create OUT folder: \"" + _OUT_FOLD + "\"")
    os.makedirs(_OUT_FOLD)
  
  
  destination_list = create_list(_SRC_LIST)
  check_duplicates (destination_list)
  for destination in destination_list:
    check_tracknumber_link (destination)
  debug_print_list (destination_list)

  create_barcode_document (destination_list, _TOP_FOLD + "/src", _OUT_FOLD + "/barcode")
  create_inserted_document (destination_list, _OUT_FOLD)

  i=0
  fold_name=_OUT_FOLD + "/envelope/"

  if not os.path.exists(fold_name): os.makedirs(fold_name)
  for file_name in os.listdir(fold_name):
    if os.path.isfile(os.path.join(fold_name, file_name)):
      os.remove(os.path.join(fold_name, file_name))

  for destination in destination_list:
    create_envelope (destination, _TOP_FOLD + "/src", fold_name, i)
    i+=1
    
  subprocess.call(f"convert {fold_name}/*.png -quality 100 {_OUT_FOLD}/envelope_full.pdf", shell=True)


if __name__ == '__main__':
  main()
