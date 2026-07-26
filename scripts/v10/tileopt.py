import os
from PIL import Image, ImageDraw, ImageFont, ImageChops
SP=os.path.dirname(os.path.abspath(__file__)); SH=os.path.join(SP,"shots")
def trim(im,pad=10):
    bb=ImageChops.difference(im.convert("L"),Image.new("L",im.size,255)).getbbox()
    if not bb: return im
    l,t,r,b=bb
    return im.crop((max(0,l-pad),max(0,t-pad),min(im.width,r+pad),min(im.height,b+pad)))
def tile(items,path,cols,cw,maxh,fs=26):
    ts=[]
    for lab,p in items:
        if not os.path.exists(p): continue
        im=trim(Image.open(p).convert("RGB")); h=max(1,int(im.height*cw/im.width))
        ts.append((lab,im.resize((cw,h),Image.LANCZOS)))
    if not ts: return
    ch=min(max(t.height for _,t in ts),maxh); lh=38
    rows=(len(ts)+cols-1)//cols
    out=Image.new("RGB",(cols*(cw+20)+20,rows*(ch+lh+20)+20),"white"); d=ImageDraw.Draw(out)
    try: f=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",fs)
    except OSError: f=ImageFont.load_default()
    for i,(l,t) in enumerate(ts):
        r,c=divmod(i,cols); x=20+c*(cw+20); y=20+r*(ch+lh+20)
        d.text((x,y+4),l,fill="black",font=f)
        box=t.crop((0,0,cw,min(t.height,ch))); out.paste(box,(x,y+lh))
        d.rectangle([x,y+lh,x+cw,y+lh+box.height],outline="#a0a0a0",width=2)
    out.save(path); print("  ",path)
L1={"A":"1A  Summary + budget side by side, platform blocks","B":"1B  One column, one squad table",
    "C":"1C  Money left, platform blocks right","D":"1D  Number strip, one squad table"}
L2={"A":"2A  Squads on top, people grouped below","B":"2B  People indented, collapsible",
    "C":"2C  Squad summary only","D":"2D  Squads on top, vacancies-only decisions"}
L3={"A":"3A  Plain tables","B":"3B  Banded column groups","C":"3C  Bridge above 3.2","D":"3D  Tiles on 3.1"}
tile([(L1[k],os.path.join(SH,f"OPT_1{k}_1_10_Z_Retail.png")) for k in "ABCD"],
     os.path.join(SH,"OPTIONS_1x.png"),2,1250,1150)
tile([(L2[k],os.path.join(SH,f"OPT_2{k}_2_3_Enterprise_Data.png")) for k in "ABCD"],
     os.path.join(SH,"OPTIONS_2x.png"),2,1300,1100)
for k in "ABCD":
    tile([("3.1 Group Summary",os.path.join(SH,f"OPT_3{k}_3_1_Group_Summary.png")),
          ("3.2 Total Cost",os.path.join(SH,f"OPT_3{k}_3_2_Total_Cost.png")),
          ("3.3 Squad Detail",os.path.join(SH,f"OPT_3{k}_3_3_Squad_Detail.png"))],
         os.path.join(SH,f"OPTIONS_3{k}.png"),3,1050,1000,22)
