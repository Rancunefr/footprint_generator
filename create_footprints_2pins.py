#!/usr/bin/env python

import os
import random
import sys
import numpy as np

from pyexcel_ods3 import get_data
from settings import *

silk_margin = pad_to_silkscreen + 0.5*line_width + solder_mask_opening

def write_header(handle, name, description):
    print( '(footprint "{}"'.format(name), file=handle )
    print( '\t(version 20240108)', file=handle )
    print( '\t(generator "footprint_generator")', file=handle )
    print( '\t(generator_version "0.1")', file=handle )
    print( '\t(layer "F.Cu")', file=handle )
    print( '\t(descr "{}")'.format(description), file=handle )
    print( '\t(attr smd)', file=handle )

def write_footer(handle):
    print( ')', file=handle )

def write_property( handle, name, value, x, y, theta, layer, hide ):
    print( '\t(property "{}" "{}"'.format(name,value), file=handle )
    print( '\t\t(at {} {} {})'.format(x,y,theta),file=handle)
    print( '\t\t(layer "{}")'.format(layer),file=handle)
    if ( hide ):
        print('\t\t(hide yes)', file=handle)
    print( '\t\t(uuid "{}")'.format( generate_uid() ),file=handle )
    print( '\t\t(effects',file=handle )
    print( '\t\t\t(font',file=handle )
    print( '\t\t\t\t(size {} {})'.format(text_size, text_size),file=handle )
    print( '\t\t\t\t(thickness {})'.format(line_width),file=handle )
    print( '\t\t\t)',file=handle )
    print( '\t\t)',file=handle )
    print( '\t)',file=handle )

def generate_uid():
    P1=random.randbytes(4) 
    P2=random.randbytes(2) 
    P3=random.randbytes(2) 
    P4=random.randbytes(2) 
    P5=random.randbytes(6) 
    result=P1.hex() + "-" + P2.hex() + "-" + P3.hex() + "-" + P4.hex() + "-" + P5.hex()
    return result

def write_pad(handle, name, x, y, size_x, size_y, bbox):
    bbox[0]=min(bbox[0], x-(size_x/2)) 
    bbox[1]=max(bbox[1], x+(size_x/2))
    bbox[2]=min(bbox[2], y-(size_y/2)) 
    bbox[3]=max(bbox[3], y+(size_y/2)) 
    print( '\t(pad "{}" smd roundrect'.format(name), file=handle )
    print( '\t\t(at {} {})'.format(x,y), file=handle )
    print( '\t\t(size {} {})'.format(size_x,size_y), file=handle )
    print( '\t\t(layers "F.Cu" "F.Paste" "F.Mask")', file=handle )
    print( '\t\t(roundrect_rratio 0.25)', file=handle )
    print( '\t\t(solder_mask_margin {})'.format(solder_mask_opening), file=handle )
    print( '\t\t(solder_paste_margin_ratio {})'.format(solder_paste_ratio), file=handle )
    print( '\t\t(thermal_bridge_angle {})'.format(thermal_bridge_angle), file=handle )
    print( '\t\t(uuid "{}")'.format( generate_uid() ),file=handle )
    print( '\t)',file=handle )

    
def write_silk(handle, type, bbox):
    if (type == "R"):
        write_silk_R( handle, bbox[0], bbox[1], bbox[2], bbox[3] )
        return
    if (type == "C"):
        write_silk_C( handle, bbox[0], bbox[1], bbox[2], bbox[3] )
        return
    if (type == "D"):
        write_silk_D( handle, bbox[0], bbox[1], bbox[2], bbox[3] )
        return
    if (type == "T"):
        write_silk_T( handle, bbox[0], bbox[1], bbox[2], bbox[3] )
        return
    return


def write_silk_R(handle, xmin, xmax, ymin, ymax):
    P_A = np.array([xmin, ymin])
    P_B = np.array([xmax, ymin])
    P_C = np.array([xmax, ymax])
    P_D = np.array([xmin, ymax])
    l = 0.2*(P_B-P_A)
    write_line( handle, "F.SilkS", P_A, P_D,   line_width )    
    write_line( handle, "F.SilkS", P_A, P_A+l, line_width )    
    write_line( handle, "F.SilkS", P_B, P_B-l, line_width )    
    write_line( handle, "F.SilkS", P_B, P_C,   line_width )    
    write_line( handle, "F.SilkS", P_C, P_C-l, line_width )    
    write_line( handle, "F.SilkS", P_D, P_D+l, line_width )    

def write_silk_C(handle, xmin, xmax, ymin, ymax):
    P_A = np.array([xmin, ymin])
    P_B = np.array([xmax, ymin])
    P_C = np.array([xmax, ymax])
    P_D = np.array([xmin, ymax])
    l = 0.2*(P_B-P_A)
    write_line( handle, "F.SilkS", P_A, P_B, line_width )    
    write_line( handle, "F.SilkS", P_B, P_C, line_width )    
    write_line( handle, "F.SilkS", P_C, P_D, line_width )    
    write_line( handle, "F.SilkS", P_D, P_A, line_width )    

def write_silk_D(handle, xmin, xmax, ymin, ymax):
    P_A = np.array([xmin, ymin])
    P_B = np.array([xmax, ymin])
    P_C = np.array([xmax, ymax])
    P_D = np.array([xmin, ymax])
    P_E = P_A - np.array([silk_margin, 0])
    P_F = P_D - np.array([silk_margin, 0])
    l = 0.2*(P_B-P_A)
    write_line( handle, "F.SilkS", P_A, P_D,   line_width )    
    write_line( handle, "F.SilkS", P_A, P_A+l, line_width )    
    write_line( handle, "F.SilkS", P_B, P_B-l, line_width )    
    write_line( handle, "F.SilkS", P_B, P_C,   line_width )    
    write_line( handle, "F.SilkS", P_C, P_C-l, line_width )    
    write_line( handle, "F.SilkS", P_D, P_D+l, line_width )    
    write_line( handle, "F.SilkS", P_E, P_F, line_width )    

def write_silk_T(handle, xmin, xmax, ymin, ymax):
    P_A = np.array([xmin, ymin])
    P_B = np.array([xmax, ymin])
    P_C = np.array([xmax, ymax])
    P_D = np.array([xmin, ymax])
    P_E = P_B + np.array([silk_margin*2, 0])
    P_F = P_C + np.array([silk_margin*2, 0])
    l = 0.2*(P_B-P_A)
    write_line( handle, "F.SilkS", P_A, P_D,   line_width )    
    write_line( handle, "F.SilkS", P_A, P_A+l, line_width )    
    write_line( handle, "F.SilkS", P_B, P_B-l, line_width )    
    write_line( handle, "F.SilkS", P_B, P_C,   line_width )    
    write_line( handle, "F.SilkS", P_C, P_C-l, line_width )    
    write_line( handle, "F.SilkS", P_D, P_D+l, line_width )    
    write_line( handle, "F.SilkS", P_E, P_F, line_width )    
    write_line( handle, "F.SilkS", P_B, P_E, line_width )    
    write_line( handle, "F.SilkS", P_C, P_F, line_width )    

def write_fab(handle, bbox):
    P_A = np.array([bbox[0], bbox[2]])
    P_B = np.array([bbox[1], bbox[2]])
    P_C = np.array([bbox[1], bbox[3]])
    P_D = np.array([bbox[0], bbox[3]])
    write_line( handle, "F.Fab", P_A, P_B, 0.05 )    
    write_line( handle, "F.Fab", P_B, P_C, 0.05 )    
    write_line( handle, "F.Fab", P_C, P_D, 0.05 )    
    write_line( handle, "F.Fab", P_D, P_A, 0.05 )    
    write_text( handle, "F.Fab", "${REFERENCE}", 0, 0, 0, 0.5, 0.08 )    

def write_fab_xy(handle, x, y):
    P_A = np.array([-x/2, -y/2])
    P_B = np.array([x/2, -y/2])
    P_C = np.array([x/2, y/2])
    P_D = np.array([-x/2, y/2])
    write_line( handle, "F.Fab", P_A, P_B, 0.05 )    
    write_line( handle, "F.Fab", P_B, P_C, 0.05 )    
    write_line( handle, "F.Fab", P_C, P_D, 0.05 )    
    write_line( handle, "F.Fab", P_D, P_A, 0.05 )    


def write_courtyard(handle, bbox):
    P_A = np.array([bbox[0], bbox[2]])
    P_B = np.array([bbox[1], bbox[2]])
    P_C = np.array([bbox[1], bbox[3]])
    P_D = np.array([bbox[0], bbox[3]])
    write_line( handle, "F.CrtYd", P_A, P_B, 0.05 )    
    write_line( handle, "F.CrtYd", P_B, P_C, 0.05 )    
    write_line( handle, "F.CrtYd", P_C, P_D, 0.05 )    
    write_line( handle, "F.CrtYd", P_D, P_A, 0.05 )    


def write_courtyard_xy(handle, x, y):
    P_A = np.array([-x/2, -y/2])
    P_B = np.array([x/2, -y/2])
    P_C = np.array([x/2, y/2])
    P_D = np.array([-x/2, y/2])
    write_line( handle, "F.CrtYd", P_A, P_B, 0.05 )    
    write_line( handle, "F.CrtYd", P_B, P_C, 0.05 )    
    write_line( handle, "F.CrtYd", P_C, P_D, 0.05 )    
    write_line( handle, "F.CrtYd", P_D, P_A, 0.05 )    


def write_line( handle, layer, P1, P2, thickness ):
    print( '\t(fp_line', file=handle )
    print( '\t\t(start {} {})'.format(P1[0],P1[1]), file=handle )
    print( '\t\t(end {} {})'.format(P2[0],P2[1]), file=handle )
    print( '\t\t(stroke', file=handle )
    print( '\t\t\t(width {})'.format(thickness), file=handle )
    print( '\t\t\t(type default)', file=handle )
    print( '\t\t)', file=handle )
    print( '\t\t(layer {})'.format(layer), file=handle )
    print( '\t\t(uuid "{}")'.format( generate_uid() ),file=handle )
    print( '\t)',file=handle )

def write_text( handle, layer, text, x, y, theta, fsize, fthick ):
    print( '\t(fp_text user "{}"'.format(text), file=handle )
    print( '\t\t(at {} {} {})'.format(x,y,theta),file=handle)
    print( '\t\t(layer "{}")'.format(layer),file=handle)
    print( '\t\t(uuid "{}")'.format( generate_uid() ),file=handle )
    print( '\t\t(effects',file=handle )
    print( '\t\t\t(font',file=handle )
    print( '\t\t\t\t(size {} {})'.format(fsize, fsize),file=handle )
    print( '\t\t\t\t(thickness {})'.format(fthick),file=handle )
    print( '\t\t\t)',file=handle )
    print( '\t\t)',file=handle )
    print( '\t)',file=handle )

def write_3dmodel( handle, model_3d, hide=False ):
    print( '\t(model "{}"'.format(model_3d), file=handle )
    if (hide):
        print('\t\t(hide yes)',file=handle)
    print( '\t\t(offset', file=handle )
    print( '\t\t\t(xyz 0 0 0)', file=handle )
    print( '\t\t)', file=handle )
    print( '\t\t(scale', file=handle )
    print( '\t\t\t(xyz 1 1 1)', file=handle )
    print( '\t\t)', file=handle )
    print( '\t\t(rotate', file=handle )
    print( '\t\t\t(xyz 0 0 0)', file=handle )
    print( '\t\t)', file=handle )
    print( '\t)', file=handle )

def expand_box( box, value ):
    return [box[0]-value, box[1]+value, box[2]-value, box[3]+value]


def main():

    if ( len(sys.argv) != 2 ):
        print("Usage: {} Lib_name".format(sys.argv[0]))
        print("")
        print("Input file:       Lib_name.ods")
        print("Output folder:    Lib_name.pretty")
        return -1

    filename = sys.argv[1]+".ods"
    folder = "./" + sys.argv[1] + ".pretty"

    if not os.path.exists(filename): 
        print("ERROR: input file not found")
        return -1

    random.seed()

    data = get_data(filename)
    sheet = data[list(data.keys())[0]]

    if not os.path.exists(folder): 
        os.makedirs(folder)

    for i, data in enumerate(sheet[1:]):
        if ( len(data) > 13 ):
            [name,description,version,type,A,B,C1,C2,D,tested,courtyard_x,courtyard_y,model_3d] = data[:13]
            alt = data[13:]
            bbox=[0,0,0,0]
            print(name)
            f = open( folder + "/" + name+".kicad_mod","w")
            write_header(f, name,description)
            pad_1="1"
            pad_2="2"
            if (type=="D"):
                pad_1="C"
                pad_2="A"
            if (type=="T"):
                pad_1="N"
                pad_2="P"
            write_pad( f, pad_1, -float(C1), 0, A, D, bbox )
            write_pad( f, pad_2,  C2, 0, B, D, bbox )
            
            if ( courtyard_x == "N/A" or courtyard_y == "N/A" ):
                pos_ref = bbox[2] - 2*silk_margin - 0.5*text_size
                pos_1 = bbox[3] + 2*silk_margin + text_size
                pos_2 = bbox[3] + 2*silk_margin + 2.5*text_size
                pos_3 = bbox[3] + 2*silk_margin + 4*text_size
                write_silk( f, type, expand_box(bbox,silk_margin) )
                write_fab( f, expand_box(bbox,silk_margin) )
                write_courtyard( f, expand_box(bbox,courtyard_margin) )
            else:
                ybox=[-courtyard_x/2, courtyard_x/2, -courtyard_y/2, courtyard_y/2]
                pos_ref = -courtyard_y/2 - 2*silk_margin - 0.5*text_size
                pos_1 = courtyard_y/2 + 2*silk_margin + text_size
                pos_2 = courtyard_y/2 + 2*silk_margin + 2.5*text_size
                pos_3 = courtyard_y/2 + 2*silk_margin + 4*text_size
                write_silk( f, type, ybox)
                write_fab( f, ybox)
                write_courtyard_xy( f, courtyard_x, courtyard_y)

            write_property(f, "Reference", "REF**", 0, pos_ref, 0, "F.SilkS", False)
            write_property(f, "Value", "Val**", 0, pos_1, 0, "F.Fab", True)
            write_property(f, "Description", description, 0, pos_2, 0, "F.Fab", True)
            write_property(f, "Is_tested", tested, 0, pos_3, 0, "F.Fab", True)
            write_property(f, "Footprint", "", 0, 0, 0, "F.Fab", True)
            write_property(f, "Datasheet", "", 0, 0, 0, "F.Fab", True)

            if ( model_3d != "N/A" ):
                write_3dmodel(f, model_3d)
            for i, alt_model in enumerate(alt):
                if (alt_model != "N/A"):
                    write_3dmodel(f, alt_model, True)
            write_footer(f)
            f.close()

if __name__ == "__main__":
    main()
