# imports
from PIL import Image, ImageDraw, ImageFont
import openpyxl
import datetime


def coupons(names: list, certificate: str, font_path: str ,font_path1: str, TODAY: str):

	for name in names:
		
		# adjust the position accordingly
		text_y_position = 1200

		# opens the image
		img = Image.open(certificate, mode ='r')
		
		# gets the image width
		image_width = img.width
		
		# gets the image height
		image_height = img.height

		# creates a drawing canvas overlay
		# on top of the image
		draw = ImageDraw.Draw(img)

		# gets the font object from the
		# font file (TTF)
		font = ImageFont.truetype(
			font_path,
			150 # change this according to your needs
		)

		# fetches the text width for
		# calculations later on
		text_width, _ = draw.textsize(name, font = font)

		draw.text(
			(
				# this calculation is done
				# to centre the image
				(image_width - text_width) / 2,
				text_y_position
			),
			name,
			font = font,
			# CMYK color code
            fill = (38,38,0,72)	 )


		font1 = ImageFont.truetype(
			font_path1,
			70 # change this according to your needs
		)
		text_width, _ = draw.textsize(TODAY, font = font1)

		draw.text(
			(
				# this calculation is done
				# to centre the image
				1000,
				2000
			),
			TODAY,
			font = font1,
			# CMYK color code
            fill = (38,38,0,72)	 )

		# saves the image in png format
		img.save("C:\\Users\\konda\\Documents\\python\\project1\\{}.png".format(name))

NAMES=[]
details_path = 'C:\\Users\\konda\\Documents\\python\\project1\\names.xlsx'
obj = openpyxl.load_workbook(details_path)
sheet = obj.active
# give range according to the names in xlsx file
for i in range(1,4):
	# Take names from Excel sheet
    get_name = sheet.cell(row = i ,column = 1)
    certi_name = get_name.value
    NAMES.append(certi_name)

# Download the font you needed and give the location of its ttf file
FONT = "C:\\Users\\konda\\Desktop\\graphics\\fonts\\Allura-Regular.ttf"

FONT1 = "C:\\Users\\konda\\Desktop\\graphics\\fonts\\Montserrat-Regular.ttf"

# path to sample certificate
CERTIFICATE = "C:\\Users\\konda\\Documents\\python\\project1\\Template.png"

#to get the date (year-month-day)
TODAY = str(datetime.date.today())
#function Call
coupons(NAMES, CERTIFICATE, FONT,FONT1,TODAY)
