#pip install pillow==9.5.0
import pandas as pd
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import textwrap
import warnings
import logging

# Configure logging
logging.basicConfig(filename='skipped_rows.log', level=logging.INFO, filemode='w',format='%(message)s')

# Ignore DeprecationWarning in the whole script
warnings.simplefilter('ignore', DeprecationWarning)

def create_barcode(data, width, height):
    # Choose the barcode format
    BARCODE_FORMAT = 'code128'

    # Create the barcode object with ImageWriter to generate an image
    barcode_obj = barcode.get_barcode_class(BARCODE_FORMAT)(data, writer=ImageWriter())

    # Generate the barcode and save it to a file-like object
    barcode_image = barcode_obj.render(writer_options={"write_text": False, 'module_width': 0.2, 'module_height': 15.0})

    # Resize the barcode to fit the specified dimensions
    barcode_image = barcode_image.resize((width, height), Image.ANTIALIAS)

    return barcode_image

def draw_multiline_text_c(draw, text, position, font, max_width, fill="black"):
    # Break the text into lines that fit within the specified width
    lines = textwrap.wrap(text, width=28)  # Adjust the 'width' parameter as needed based on your font and desired line length
    y = position[1]
    for line in lines:
        line_width, line_height = draw.textsize(line, font=font)
        # Center align the text
        x = position[0] + (max_width - line_width) / 2
        draw.text((x, y), line, font=font, fill=fill)
        y += line_height+5  # Move to the next line

def draw_multiline_text_l(draw, text, position, font, max_width, fill="black"):
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        # Check width of the line with the new word added
        test_line = current_line + " " + word if current_line else word
        width, _ = draw.textsize(test_line, font=font)
        if width <= max_width:
            # If the line with the new word fits within the max width, update the current line
            current_line = test_line
        else:
            # If the new word would exceed the max width, start a new line
            lines.append(current_line)
            current_line = word
    lines.append(current_line)  # Add the last line

    y = position[1]
    for line in lines:
        draw.text((position[0], y), line, font=font, fill=fill)
        draw.text((position[0]+0.3, y+0.3), line, font=font, fill=fill)
        draw.text((position[0]-0.3, y+0.3), line, font=font, fill=fill)
        _, line_height = draw.textsize(line, font=font)
        y += line_height

def icardfront(name, new_roll, old_roll, branch,validity, image_path, front_output_path):
    # Load base ID card template
    id_card = Image.open("card1.png")  

    # Load user's image
    user_image = Image.open(image_path)  # Replace with the actual path to user's image

    # Resize user's image to fit the ID card template
    box_width, box_height = 392 - 201, 448 - 264
    user_image = user_image.resize((box_width, box_height))

    # Paste user's image onto the ID card template at the specified coordinates
    id_card.paste(user_image, (201, 264))

    # Set up drawing context
    draw = ImageDraw.Draw(id_card)

    # Load font
    font = ImageFont.truetype("arial.ttf", 36)
    max_name_width = 570-45
    if(len(name)>21):
        font = ImageFont.truetype("arial.ttf", 32)
        draw_multiline_text_c(draw, name.upper(), (43, 475), font, max_name_width,"white")
    else :
        draw_multiline_text_c(draw, name.upper(), (43, 488), font, max_name_width,"white")

    # Calculate the width of the name text to center-align it
    # name_width, _ = font.getsize(name.upper())
    # name_start_pos = (id_card.width - name_width) // 2  # Center-align the text
    # # Write user's name on ID card, center-aligned
    # draw.text((name_start_pos, 488), name.upper(), fill="white", font=font)

    # Write user's New Roll and department on ID card
    font = ImageFont.truetype("arial.ttf", 41)
    id_width, _ = font.getsize("DTU/" + old_roll)
    id_start_pos = (id_card.width - id_width) // 2 
    draw.text((id_start_pos, 558), "DTU/" + old_roll, fill="black", font=font)
    draw.text((id_start_pos+0.3, 558+0.3), "DTU/" + old_roll, fill="black", font=font)
    draw.text((id_start_pos-0.3, 558-0.3), "DTU/" + old_roll, fill="black", font=font)

    # Write user's old Roll and department on ID card
    font = ImageFont.truetype("arial.ttf", 34)
    id_width, _ = font.getsize(new_roll)
    id_start_pos = (id_card.width - id_width) // 2 
    draw.text((id_start_pos, 619),new_roll, fill="black", font=font)
    draw.text((id_start_pos+0.3, 619+0.3), new_roll, fill="black", font=font)
    draw.text((id_start_pos-0.3, 619-0.3), new_roll, fill="black", font=font)

    # Write user's old Roll and department on ID card
    font = ImageFont.truetype("arial.ttf", 34)
    id_width, _ = font.getsize(branch)
    id_start_pos = (id_card.width - id_width) // 2 
    draw.text((id_start_pos, 658),branch, fill="black", font=font)
    draw.text((id_start_pos+0.3, 658+0.3), branch, fill="black", font=font)
    draw.text((id_start_pos-0.3, 658-0.3), branch, fill="black", font=font)

    # Write Valid Upto on ID card
    font = ImageFont.truetype("arial.ttf", 34)
    id_width, _ = font.getsize("Valid Upto : Jul 2022 to " + validity)
    id_start_pos = (id_card.width - id_width) // 2 
    draw.text((id_start_pos, 706),"Valid Upto : Jul 2022 to " + validity, fill="black", font=font)
    draw.text((id_start_pos+0.3, 706+0.3), "Valid Upto : Jul 2022 to " + validity, fill="black", font=font)
    draw.text((id_start_pos-0.3, 706-0.3), "Valid Upto : Jul 2022 to " + validity, fill="black", font=font)

    # Save the generated ID card
    id_card.save(front_output_path)

def icardback(pname, mobno, email, codeid,address, back_output_path):
    # Load base ID card template
    id_card = Image.open("card2.png")  # Ensure this path points to your ID card template

    # Resize user's image to fit the ID card template
    box_width, box_height = 468 - 145, 519 - 441
    # Resize user's image to fit the specified box
    barcode_img = create_barcode(codeid, box_width, box_height)
    # Paste user's image onto the ID card template at the top-left corner's coordinates
    id_card.paste(barcode_img, (145, 441))

    # Set up drawing context
    draw = ImageDraw.Draw(id_card)

    if(len(pname) > 18):
        font = ImageFont.truetype("arial.ttf", 28)
        max_pname_width = 570 - 40
        draw_multiline_text_c(draw, "Parent: " + pname.upper(), (40, 29), font, max_pname_width)
        
    else: 
        # Write user's Parent's Name on ID card
        font = ImageFont.truetype("arial.ttf", 32)
        max_pname_width = 570 - 40
        draw_multiline_text_c(draw, "Parent: " + pname.upper(), (40, 29), font, max_pname_width)

    # Write user's MobileNo. on ID card
    font = ImageFont.truetype("arial.ttf", 34)
    id_width, _ = font.getsize("Mob: " + mobno)
    id_start_pos = (id_card.width - id_width) // 2 
    draw.text((id_start_pos, 240),"Mob: " + mobno, fill="black", font=font)
    draw.text((id_start_pos+0.3, 240+0.3), "Mob: " + mobno, fill="black", font=font)
    draw.text((id_start_pos-0.3, 240-0.3), "Mob: " + mobno, fill="black", font=font)

    # Write user's email on ID card
    font = ImageFont.truetype("arial.ttf", 34)
    max_email_width = 570 - 40
    draw_multiline_text_c(draw, email, (40, 300), font, max_email_width)
    
    # Write user's codeid on ID card
    font = ImageFont.truetype("arial.ttf", 34)
    id_width, _ = font.getsize(codeid)
    draw.text((220, 531),codeid, fill="black", font=font)
    draw.text((220+0.3, 531+0.3), codeid, fill="black", font=font)
    draw.text((220-0.3, 531-0.3), codeid, fill="black", font=font)

    # Define font for the address
    font = ImageFont.truetype("arial.ttf", 25)
    # Define the maximum width for the address
    max_address_width = 580 - 35  # x=543 - x=35
    # Write user's address, ensuring it doesn't cross x=543
    draw_multiline_text_l(draw, address, (35, 92), font, max_address_width)

    # Save the generated ID card
    id_card.save(back_output_path)

def format_date(date_str):
    # Parse the date string into a datetime object
    # Adjust format to match the structure "YYYY-MM-DD HH:MM:SS"
    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
    # Format the datetime object into the desired string format "Month YYYY"
    formatted_date = date_obj.strftime('%b %Y')
    return formatted_date


# Read data from the Excel file
data_df = pd.read_excel('data.xlsx')
# Format the 'validity' column with the correct datetime format
data_df['validity'] = data_df['validity'].astype(str).apply(format_date)

# Loop over each row in the DataFrame
for index, row in data_df.iterrows():
    # Extract data for the current student
    name = str(row['name'])
    old_roll = str(row['old_rollno'])
    new_roll = str(row['new_rollno'])
    branch = str(row['branch'])
    mobno = str(row['mobileno'])
    email = str(row['email'])
    pname = str(row['parentname'])
    address = str(row['address'])
    codeid = str(row['codeid'])
    image_path = str(row['imagepath'])
    validity = str(row['validity'])
    # bcodeimage_path = str(row['bcodeimage'])

    # Error Counter
    error = 0

    # Check if name length is greater than 21 characters
    if len(pname) > 42:
        error = 1
        # Log the data for this row
        logging.info(f'{new_roll} ||Parent Name length Exceeded (>42) : here {len(pname)}')
        # Skip the rest of the loop and move to the next iteration
        
    if len(address) > 152:
        error = 1
        # Log the data for this row
        logging.info(f'{new_roll} || Address length Exceeded (>152) : here {len(address)}')
        # Skip the rest of the loop and move to the next iteration
        
    if len(name) > 48:
        error = 1
        # Log the data for this row
        logging.info(f'{new_roll} || Name length Exceeded (>48): here {len(name)}')
        # Skip the rest of the loop and move to the next iteration
        
    if len(email) > 53:
        error = 1
        # Log the data for this row
        logging.info(f'{new_roll} || Email length Exceeded (>53): here {len(email)}')
        # Skip the rest of the loop and move to the next iteration
        
    if len(mobno) > 19:
        error = 1
        # Log the data for this row
        logging.info(f'{new_roll} || Mobile length Exceeded (>19): here {len(mobno)}')
        # Skip the rest of the loop and move to the next iteration
    if len(new_roll) < 6 or len(old_roll)< 6:
        error = 1
        # Log the data for this row
        logging.info(f'{new_roll} -- {old_roll} || Invalid Roll number')
        # Skip the rest of the loop and move to the next iteration
        
    if(error != 0):
        continue

    # Generate file names for the output ID cards
    front_output_path = new_roll.replace("/", "_")+"a.png"
    back_output_path = new_roll.replace("/", "_")+"b.png"
    
    # Call functions to generate the front and back of the ID card
    icardfront(name, new_roll, old_roll, branch, validity, "studentphoto/" + image_path, "output/"+front_output_path)
    icardback(pname, mobno, email, codeid, address, "output/"+ back_output_path)

    front_image = Image.open("output/"+front_output_path)
    back_image = Image.open("output/"+back_output_path)

    # Ensure images are in RGB mode for PDF conversion
    front_image = front_image.convert("RGB")
    back_image = back_image.convert("RGB")
    
    front_image.save("pdf/"+new_roll.replace("/","_")+".pdf", "PDF", resolution=400, save_all=True, append_images=[back_image],quality=100)

    