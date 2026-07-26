import os,re
from PIL import Image, ImageDraw, ImageFont, ImageChops
SP=os.path.dirname(os.path.abspath(__file__)); SH=os.path.join(SP,"shots")
LAB={"1A":"1A  Stacked blocks","1B":"1B  Two columns at the top","1C":"1C  Platform as a column",
     "1D":"1D  Number strip on top","2A":"2A  Squads on top, people grouped below",
     "2B":"2B  One table, people indented and collapsible","2C":"2C  Squads only, people on one FTE tab",
     "2D":"2D  Squads on top, vacancies-only decision table","3A":"3A  One table per tab",
     "3B":"3B  Column groups: Design / Today / After","3C":"3C  Vertical bridge","3D":"3D  Tiles then table"}
def trim(im,pad=10):
    g=im.convert("L"); bg=Image.new("L",im.size,255)
    bb=ImageChops.difference(g,bg).getbbox()
    if not bb: return im
    l,t,r,b=bb
    return im.crop((max(0,l-pad),max(0,t-pad),min(im.width,r+pad),min(im.height,b+pad)))
for fam in "123":
    keys=[f"{fam}{x}" for x in "ABCD"]
    tiles=[]
    for k in keys:
        p=os.path.join(SH,f"opt_{k}.png")
        if not os.path.exists(p): continue
        im=trim(Image.open(p).convert("RGB"))
        w=1150; h=max(1,int(im.height*w/im.width))
        tiles.append((k,im.resize((w,h),Image.LANCZOS)))
    if not tiles: continue
    ch=min(max(t.height for _,t in tiles),950); lab=40
    W=2*(1150+20)+20; H=2*(ch+lab+20)+20
    out=Image.new("RGB",(W,H),"white"); d=ImageDraw.Draw(out)
    try: f=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",26)
    except OSError: f=ImageFont.load_default()
    for i,(k,t) in enumerate(tiles):
        r,c=divmod(i,2); x=20+c*(1150+20); y=20+r*(ch+lab+20)
        d.text((x,y+6),LAB[k],fill="black",font=f)
        box=t.crop((0,0,1150,min(t.height,ch)))
        out.paste(box,(x,y+lab))
        d.rectangle([x,y+lab,x+1150,y+lab+box.height],outline="#b0b0b0",width=2)
    p=os.path.join(SH,f"options_{fam}x.png"); out.save(p); print("  ",p)
