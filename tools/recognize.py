from base64 import b64decode
import datetime, io
from PIL import Image

dir = 'captchas'

# [4x, 1, 2, 3, 4, 5, 6, 7, 8, 9]
maskList = [
    "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"]


def saveCaptchaImg(CaptchaImg, fn):
    global dir
    i = 0
    DuplicateNamesFound = False
    while not DuplicateNamesFound:
        fnbase = "captcha " + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + " = " + fn
        try:
            if i == 0:
                fullfn = fnbase + ".png"
            else:
                fullfn = fnbase + f"({i})" + ".png"
            CaptchaImg.save(fullfn)
            DuplicateNamesFound = True
        except FileExistsError:
            i = i + 1
            DuplicateNamesFound = False
        except:
            break


def recognize(imgData, saveImg=False):
    Zlist = []  # [(x1, z1), (x2, z2), (x3, z3), etc.] - position and digit
    captcha = ""
    global maskList
    originalimage = Image.open(io.BytesIO(b64decode(imgData)))
    processedimage = originalimage.convert('L').point(lambda x: 255 if x > 20 else 0, mode='1').convert('1').convert(
        'RGBA')
    if originalimage.getextrema() == ((0, 0), (0, 0), (0, 0), (255, 255)):
        return ("empty image")
    for z in [4, 2, 3, 1, 5, 6, 7, 8, 9]:  # reorder to exclude false 1 on 4
        mask = Image.open(io.BytesIO(b64decode(maskList[z])))
        previ = 0
        for i in range(15, 120):  # leftmost mask position, rigthmost mask position
            resultimage = Image.alpha_composite(processedimage.crop((i, 0, i + 21, 0 + 50)), mask)
            if resultimage.getextrema() == ((0, 0), (0, 0), (0, 0), (255, 255)):
                if z == 4:  # fill 4 with white to exclude false 1 on 4
                    maskx = Image.open(io.BytesIO(b64decode(maskList[0])))
                    processedimage.paste(Image.alpha_composite(processedimage.crop((i, 0, i + 21, 0 + 50)), maskx),
                                         (i, 0))
                if previ == 0 or i > previ + 15:  # no digit closer then 15 px
                    Zlist.append((i, z))
                    if len(Zlist) == 5:
                        Zlist.sort()
                        for z in Zlist:
                            captcha = captcha + str(z[1])
                        if saveImg:
                            saveCaptchaImg(originalimage, captcha)
                        return (captcha)
                    previ = i
                    i = i + 15  # skip a little
    Zlist.sort()
    if saveImg:
        saveCaptchaImg(originalimage, captcha)
    return (str(Zlist))  # if less then 5 digits recognized
