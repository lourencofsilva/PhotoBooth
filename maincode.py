#!/usr/bin/env python
import PySimpleGUI as sg
import os
from PIL import Image, ImageTk
import io
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import shutil
import cv2
import ctypes
user32 = ctypes.windll.user32

height = user32.GetSystemMetrics(1)
width = user32.GetSystemMetrics(0)

update = False


def getgoodfontvalue(originalvalue):
    originaldisplayvalue = 1536
    number = originaldisplayvalue / originalvalue
    thispcnumber = width / number
    return int(thispcnumber)


# Get the folder containing the images from the user

iconlocation = b"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"
cliplogo = b"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"

resourcesfolder = sg.popup_get_folder("Resources Folder", default_path="", icon=iconlocation)
if not resourcesfolder:
    sg.popup_cancel()
    raise SystemExit()

folder = sg.popup_get_folder('700D Images Folder', default_path='', icon=iconlocation)
if not folder:
    sg.popup_cancel()
    raise SystemExit()

tempfolder = sg.popup_get_folder('Modified Photos Folder', default_path='', icon=iconlocation)
if not tempfolder:
    sg.popup_cancel()
    raise SystemExit()

# PIL supported image types
img_types = (".png", ".jpg", "jpeg", ".tiff", ".bmp")

while True:
    # get list of files in folder
    flist0 = os.listdir(folder)

    # create sub list of image files (no sub folders, no wrong file types)
    fnames = [f for f in flist0 if os.path.isfile(
        os.path.join(folder, f)) and f.lower().endswith(img_types)]

    num_files = len(fnames)  # number of images found
    if num_files == 0:
        sg.popup('No files in folder')
        raise SystemExit()

    for i10 in fnames:
        i10filename = os.path.join(folder, i10)
        if i10filename.endswith('.JPG'):
            i10newfilename = os.path.splitext(i10filename)[0]
            im = Image.open(i10filename)
            im.save(i10newfilename + ".png")
            os.remove(i10filename)

    # get list of files in folder
    flist0 = os.listdir(folder)

    # create sub list of image files (no sub folders, no wrong file types)
    fnames = [f for f in flist0 if os.path.isfile(
        os.path.join(folder, f)) and f.lower().endswith(img_types)]

    for i20 in fnames:
        i20filename = os.path.join(folder, i20)
        i20newfilename = os.path.join(tempfolder, i20)
        if not os.path.exists(i20newfilename):
            shutil.copy(i20filename, tempfolder)

    del flist0  # no longer needed


    # ------------------------------------------------------------------------------
    # use PIL to read data of one image
    # ------------------------------------------------------------------------------

    def get_img_data(f, maxsize=(getgoodfontvalue(1000), getgoodfontvalue(500)), first=False):
        """Generate image data using PIL
        """
        img = Image.open(f)
        img.thumbnail(maxsize)
        if first:  # tkinter is inactive the first time
            bio = io.BytesIO()
            img.save(bio, format="PNG")
            del img
            return bio.getvalue()
        return ImageTk.PhotoImage(img)


    # initialize to the first file in the list
    originalimage = True
    filename = os.path.join(folder, fnames[-1])  # name of first file in list
    image_elem = sg.Image(data=get_img_data(filename, first=True))
    filenamestring = "Name of Picture File: " + fnames[-1]
    filename_display_elem = sg.Text(filenamestring, size=(38, 1), font=("Arial", getgoodfontvalue(20)))
    currenttextelement = sg.Text("Seeing Original Photo", font=("Arial", getgoodfontvalue(30)), text_color="red", auto_size_text=True)
    file_num_display_elem = sg.Text('File {} of {}'.format(num_files, num_files), size=(15, 1))

    # define layout, show and read the form
    col = [[filename_display_elem, sg.Button("Update Photos", font=("Arial", getgoodfontvalue(20)))],
           [image_elem],
           [sg.Button("Original Photo", size=(12, 1), font=("Arial", getgoodfontvalue(20))), currenttextelement, sg.Button("Modified Photo", size=(12, 1), font=("Arial", getgoodfontvalue(20)))],
           [sg.Button("Reset Picture"), sg.Button("B&W"), sg.Button("Filter 1"), sg.Button("Filter 2"), sg.Button("Filter 3"), sg.Button("Filter 4"), sg.Button("Filter 5"), sg.Button("Filter 6"), sg.Button("Filter 7"), sg.Button("Filter 8"), sg.Button("Border 1"), sg.Button("Border 2"), sg.Button("Border 3"), sg.Button("Border 4")]]

    col_files = [[sg.Button(image_data=cliplogo, button_color=sg.COLOR_SYSTEM_DEFAULT, border_width=0)],
                 [sg.Listbox(values=fnames, change_submits=True, size=(40, 20), key='listbox')],
                 [sg.Button('Prev', size=(8, 2)), sg.Button('Next', size=(8, 2)), file_num_display_elem],
                 [sg.Text()],
                 [sg.Button("Print Photo", size=(8, 2), font=("Arial", getgoodfontvalue(20))),
                  sg.Button("E-Mail Photo", size=(8, 2), font=("Arial", getgoodfontvalue(20)))]]

    layout = [[sg.Column(col_files), sg.Column(col)]]

    window = sg.Window('CLIP Media Club Photo Booth', layout, return_keyboard_events=True,
                       location=(0, 0), use_default_focus=False, icon=iconlocation)

    # loop reading the user input and displaying image, filename
    i = num_files - 1
    while True:
        # read the form
        event, values = window.read()
        # perform button and keyboard operations
        if event is None:
            exit(0)
        elif event in ('Next', 'MouseWheel:Down', 'Down:40', 'Next:34'):
            i += 1
            if i >= num_files:
                i -= 1
            filename = os.path.join(folder, fnames[i])
            filenamestring = "Name of Picture File: " + fnames[i]
        elif event in ('Prev', 'MouseWheel:Up', 'Up:38', 'Prior:33'):
            i -= 1
            if i < 0:
                i += 1
            filename = os.path.join(folder, fnames[i])
            filenamestring = "Name of Picture File: " + fnames[i]
        elif event == 'listbox':  # something from the listbox
            f = values["listbox"][0]  # selected filename
            filename = os.path.join(folder, f)  # read this file
            i = fnames.index(f)  # update running index
        elif event in "Update Photos":
            break
        elif event in "Print Photo":
            os.startfile(filename, "print")
        elif event in "E-Mail Photo":
            auto_update = False
            emailaddress = sg.PopupGetText("Enter E-Mail to send photo:", icon=iconlocation)
            if emailaddress is None or emailaddress == "":
                sendemail = False
            else:
                sendemail = True

            if sendemail:
                layoutaskforemailaddress = [[sg.Text("Please wait while E-Mail is sent!")]]
                askforemailaddresswindow = sg.Window("E-Mail Sending", layoutaskforemailaddress, icon=iconlocation)
                img_data = open(filename, 'rb').read()
                msg = MIMEMultipart()
                msg['Subject'] = "EMAIL SUBJECT"
                msg['From'] = "EMAIL ADDRESS"
                msg['To'] = emailaddress

                text = MIMEText("""EMAIL CONTENT""")
                msg.attach(text)
                with open(filename, "rb") as file:
                    part = MIMEApplication(
                        file.read(),
                        Name=basename(filename)
                    )
                # After the file is closed
                part['Content-Disposition'] = 'attachment; filename="%s"' % basename(filename)
                msg.attach(part)
                s = smtplib.SMTP("EMAIL SERVER", 587)
                s.ehlo()
                s.starttls()
                s.ehlo()
                s.login("EMAIL ADDRESS", "EMAIL PASSWORD")
                s.sendmail("EMAIL ADDRESS", emailaddress, msg.as_string())
                s.quit()
                successpopupstring = "E-Mail sent successfully to: " + emailaddress
                askforemailaddresswindow.Close()
                sg.Popup(successpopupstring, icon=iconlocation)
            auto_update = True
        elif event in "Original Photo":
            originalimage = True
        elif event in "Modified Photo":
            originalimage = False
        elif event in "B&W":
            filename = os.path.join(tempfolder, fnames[i])
            img = Image.open(filename).convert("LA")
            img.save(filename)
        elif event in "Filter 1":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-1.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Filter 2":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-2.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Filter 3":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-3.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Filter 4":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-4.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Filter 5":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-5.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Filter 6":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-6.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Filter 7":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-7.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Filter 8":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Filters/filter-8.png"
            color_img = cv2.imread(color_img_filename)
            target_image = cv2.imread(filename)
            fused_img = cv2.addWeighted(target_image, 0.9, color_img, 0.25, 0)
            cv2.imwrite(filename, fused_img)
        elif event in "Border 1":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Borders/border-1.png"
            img = Image.open(color_img_filename)
            background = Image.open(filename)
            background.paste(img, (0, 0), img)
            background.save(filename, format="png")
        elif event in "Border 2":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Borders/border-2.png"
            img = Image.open(color_img_filename)
            background = Image.open(filename)
            background.paste(img, (0, 0), img)
            background.save(filename, format="png")
        elif event in "Border 3":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Borders/border-3.png"
            img = Image.open(color_img_filename)
            background = Image.open(filename)
            background.paste(img, (0, 0), img)
            background.save(filename, format="png")
        elif event in "Border 4":
            filename = os.path.join(tempfolder, fnames[i])
            color_img_filename = resourcesfolder + "/Borders/border-4.png"
            img = Image.open(color_img_filename)
            background = Image.open(filename)
            background.paste(img, (0, 0), img)
            background.save(filename, format="png")
        elif event in "Reset Picture":
            os.remove(filename)
            originalfile = os.path.join(folder, fnames[i])
            modifiedfile = os.path.join(tempfolder, fnames[i])
            shutil.copy(originalfile, modifiedfile)
        else:
            filename = os.path.join(folder, fnames[i])
            filenamestring = "Name of Picture File: " + fnames[i]

        # update window with new image
        if originalimage:
            currenttextelement.update("Seeing Original Photo")
            filename = os.path.join(folder, fnames[i])
        else:
            currenttextelement.update("Seeing Modified Photo")
            filename = os.path.join(tempfolder, fnames[i])
        image_elem.update(data=get_img_data(filename, first=True))
        # update window with filename
        filename_display_elem.update(filenamestring)
        # update page display
        file_num_display_elem.update('File {} of {}'.format(i + 1, num_files))

    window.close()
