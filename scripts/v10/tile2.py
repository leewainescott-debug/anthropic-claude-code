import os,re
from PIL import Image, ImageDraw, ImageFont, ImageChops
SP=os.path.dirname(os.path.abspath(__file__)); SH=os.path.join(SP,"shots")
def trim(im,pad=10):
    g=im.convert("L"); bb=ImageChops.difference(g,Image.new("L",im.size,255)).getbbox()
    if not bb: return im
    l,t,r,b=bb
    return im.crop((max(0,l-pad),max(0,t-pad),min(im.width,r+pad),min(im.height,b+pad)))
def tile(items,path,cols,cw,maxh,fs=24):
    tiles=[]
    for lab,p in items:
        if not os.path.exists(p): continue
        im=trim(Image.open(p).convert("RGB"))
        h=max(1,int(im.height*cw/im.width)); tiles.append((lab,im.resize((cw,h),Image.LANCZOS)))
    if not tiles: return
    ch=min(max(t.height for _,t in tiles),maxh); lab_h=36
    rows=(len(tiles)+cols-1)//cols
    out=Image.new("RGB",(cols*(cw+20)+20, rows*(ch+lab_h+20)+20),"white"); d=ImageDraw.Draw(out)
    try: f=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",fs)
    except OSError: f=ImageFont.load_default()
    for i,(l,t) in enumerate(tiles):
        r,c=divmod(i,cols); x=20+c*(cw+20); y=20+r*(ch+lab_h+20)
        d.text((x,y+4),l,fill="black",font=f)
        box=t.crop((0,0,cw,min(t.height,ch))); out.paste(box,(x,y+lab_h))
        d.rectangle([x,y+lab_h,x+cw,y+lab_h+box.height],outline="#b0b0b0",width=2)
    out.save(path); print("  ",path)
L1={"1A":"1A  Owner's arrangement, cleaned","1B":"1B  Stacked, one squad table",
    "1C":"1C  Money left, squads right","1D":"1D  Number strip, one squad table"}
tile([(L1["1"+k],os.path.join(SH,f"o2_1{k}_1_1_Ampol_Retail.png")) for k in "ABCD"],
     os.path.join(SH,"opts2_1x.png"),2,1250,1000)
L3={"A":"3A  Plain tables","B":"3B  Column groups","C":"3C  Bridge on 3.2","D":"3D  Tiles on 3.1"}
for k in "ABCD":
    tile([(f"3.1 Group Summary",os.path.join(SH,f"o2_3{k}_3_1_Group_Summary.png")),
          (f"3.2 Total Cost",os.path.join(SH,f"o2_3{k}_3_2_Total_Cost.png")),
          (f"3.3 Squad Detail",os.path.join(SH,f"o2_3{k}_3_3_Squad_Detail.png"))],
         os.path.join(SH,f"opts2_3{k}.png"),3,1000,900,22)
