#importing libraries
import win32com.client 
import os
from pdf2image import convert_from_path
import urllib.request
import pyttsx3
from gtts import gTTS
from moviepy.editor import *
import shutil
import PyPDF2
import time
from sys import exit
import fitz
import glob
import pythoncom
#pip install PyMuPDF too

def mainw(filepath):
    # in_file=input("enter the path: ")
    pythoncom.CoInitialize()

    in_file = f"ppt2vid\media\\{filepath}"
    out_file = os.path.splitext(in_file)[0]
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    pdf = powerpoint.Presentations.Open(in_file,WithWindow = False)
    pdf.SaveAs(out_file,32)
    pdf.Close()  
    powerpoint.Quit()
    pythoncom.CoUninitialize()

    dir =os.path.normpath(in_file+ os.sep + os.pardir)
    #creating a temp folder
    os.mkdir(dir+"\\temp")
    temp =dir+"\\temp"
    os.mkdir(dir+"\\temp"+"\\images")
    images_folder=dir+"\\temp"+"\\images"
    os.mkdir(dir+"\\temp"+"\\videos")
    videos_folder=dir+"\\temp"+"\\videos"
    os.mkdir(dir+"\\temp"+"\\speech")
    speech_folder=dir+"\\temp"+"\\speech"

    file = open(out_file+".pdf", 'rb')
    readpdf = PyPDF2.PdfFileReader(file)
    totalpages = readpdf.numPages

    doc = fitz.open(out_file + '.pdf')
    zoom = 4
    mat = fitz.Matrix(zoom, zoom)

    ppt_dir = in_file
    ppt_app = win32com.client.GetObject(ppt_dir)
    listi = []
    for i in range(totalpages):
        listi.append(i)
    for i in range(totalpages):
        val = images_folder+f"/{i}.jpg"
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=mat)
        pix.save(val)
    doc.close()

    #below function contains all process which is sonverting our images to video and comments to voice over and merging to create a final file
    def important():
            for i,ppt_slide in zip(listi,ppt_app.Slides):
                #checking if slide contain comment or not
                if(ppt_slide.Comments.Count == 0):
                    #checking if comment in empty
                            clip = ImageClip(images_folder+f'\{i}.jpg').set_duration(3)
                            clip.write_videofile(videos_folder+f"\{i}.mp4",fps=24,remove_temp= True,codec="libx264",audio_codec='aac')
                else:
                    for comment in ppt_slide.Comments:
                        # print(ppt_slide.Comments.Count)
                        if(len(comment.Text)==0):  
                            clip = ImageClip(images_folder+f'\{i}.jpg').set_duration(3)
                            clip.write_videofile(videos_folder+f"\{i}.mp4",fps=24,remove_temp= True,codec="libx264",audio_codec='aac')

                        else:
                            #checking wheather device is connected with internet or not
                            def connect():
                                try:
                                    urllib.request.urlopen('http://google.com') #Python 3.x
                                    return True
                                except:
                                    return False
                            if connect() == True:
                                mytext = comment.Text
                                language = 'en'
                                myobj = gTTS(text=mytext, lang=language, slow=False,tld='co.in')
                                myobj.save(speech_folder+f'\{i}.mp3')
                            else:
                                engine = pyttsx3.init()
                                text = comment.Text
                                engine.say(text)
                                engine.save_to_file(text,speech_folder+f'\{i}.mp3')
                                engine.runAndWait()
                            audio = AudioFileClip(speech_folder+f'\{i}.mp3')
                            clip = ImageClip(images_folder+f'\{i}.jpg').set_duration(audio.duration)
                            clip = clip.set_audio(audio) 
                            clip.write_videofile(videos_folder+f"\{i}.mp4", fps=24)
            #merging all files 
            clips=[]
            files = glob.glob(os.path.expanduser(videos_folder+'\*'))
            sorted_by_mtime_ascending = sorted(files, key=lambda t: os.stat(t).st_mtime)
            for i in sorted_by_mtime_ascending:
                print(i)
                video = VideoFileClip(i)
                clips.append(video)
            final = concatenate_videoclips(clips,method='compose')
            final.write_videofile(dir+"\\final.mp4",fps=24,remove_temp= True,codec="libx264")
            return None

    def final_del():
            shutil.rmtree(temp)

            
    important()
    #deleteing temp files generated
    final_del()


