from PIL import Image, ImageDraw, ImageFont
import os, shutil, pandas

# Global Variables
spreadsheet_file = 'tedxparticipants.xlsx'


# fetches list of names from spreadsheet
def getNames():
    workbook=pandas.read_excel(spreadsheet_file)
    nameList=workbook['Name'].tolist()
    return nameList

# fetches list of positions from spreadsheet
def getPositions():
    workbook=pandas.read_excel(spreadsheet_file)
    PosList=workbook['Position'].tolist()
    return PosList


# Creates a new 'generated' folder if not already present
def folder_check():
    if os.path.isdir("generated_certificates"): 
        shutil.rmtree("generated_certificates")
        os.mkdir("generated_certificates")
    else:
        os.mkdir("generated_certificates")
     
# certificate generation via editing the template
def generate(names,pos):
    folder_check()
    for i in range(len(names)):  
        print("Generating " + names[i]+"@ " + pos[i],end=' ')
        
        if pos[i].find("Team")!=-1:
            img = Image.open("templates/vol_temp.png") # Loading the certificate template
            print("Volunteer")
        else:
            print("Lead")
            img = Image.open("templates/lead_temp.png") # Loading the certificate template
        draw = ImageDraw.Draw(img)

        font = ImageFont.truetype('fonts/PlayfairDisplay-SemiBold.ttf', 250) # Setting the font to Poppins Medium and font size to 40
        w, h = draw.textsize(names[i], font=font)
        print("Coordinates: ",w,h)    
        name_x, name_y = (1280-w)/2, (((82-h)/2)+434) # Setting the co-ordinates to where the names should be entered
        name_x, name_y = 100,2275
        draw.text((name_x, name_y), names[i], font=font, fill="black")

        #
        font = ImageFont.truetype('fonts/PlayfairDisplay-SemiBold.ttf', 200) # Setting the font to Poppins Medium and font size to 40
        w, h = draw.textsize(pos[i], font=font)   
        name_x, name_y = (1280-w)/2, (((82-h)/2)+434) # Setting the co-ordinates to where the names should be entered
        name_x, name_y =200,2805
        draw.text((name_x, name_y), pos[i], font=font, fill="black")
        #
                                                                                                                               
        img.save(r'generated_certificates/' + names[i] + ".png") # Saving the images to the generated_certificates directory


names = getNames()
pos= getPositions()

generate(names,pos)
