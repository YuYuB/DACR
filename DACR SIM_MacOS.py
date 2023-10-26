
import random
import os

import math
import xlwt
from xlwt import Workbook

test=Workbook(encoding='ascii')


feuil1=test.add_sheet("X")

feuil2=test.add_sheet("A")

style = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;' 'font: colour white, bold True;' 'align: horiz center,vert center;')
style4 = xlwt.easyxf('pattern: pattern solid, fore_colour black;' 'font: colour yellow;' 'align:horiz center,vert center;')
style5 = xlwt.easyxf('pattern: pattern solid, fore_colour red;' 'font: bold True;' 'align:horiz center,vert center;')
        
stylebis = xlwt.easyxf('pattern: pattern solid, fore_colour red;' 'font: colour white, bold True;' 'align:horiz center,vert center;')
styleter=xlwt.easyxf("pattern: pattern solid, fore_color green; font: color white; align: horiz center")


feuil1.row(0).height = 516
feuil1.col(0).width = 4000
feuil1.col(1).width = 5000
feuil1.col(2).width = 5000
feuil1.col(3).width = 5000
feuil1.col(4).width =   5000
feuil1.col(5).width = 5000

feuil1.col(6).width = 1000
feuil1.col(7).width = 5000
feuil1.col(8).width = 5000
feuil1.col(9).width = 5000
feuil1.col(10).width = 5000
feuil1.col(10).width = 5000
feuil1.col(11).width = 7000



feuil2.row(0).height = 200


feuil1.row(0).height = 200


feuil2.col(0).width = 4000
feuil2.col(1).width = 5000
feuil2.col(2).width = 2000
feuil2.col(3).width = 5000
feuil2.col(4).width = 5000
feuil2.col(5).width = 500

feuil2.col(6).width = 500
feuil2.col(7).width = 5000
feuil2.col(8).width = 5000

feuil2.col(9).width = 5000
feuil2.col(10).width = 5000
feuil2.col(11).width = 7000

feuil1.write(0,1,'n',style)
#feuil1.write(0,9,"event",style)
feuil2.write(0,1,'n',style)
#feuil2.write(0,9,"event",style)
feuil2.write(0,3,"Test A (onset)",style)
feuil1.write(0,2,"Test X (onset)",style)
feuil1.write(0,8,"outcome",style)
feuil2.write(0,8,"outcome",style)

feuil1.write(0,4,'Event',style)
feuil1.write(0,5,'nX+',style)
feuil1.write(0,10,"exponent of It0 for X",style)
feuil2.write(0,4,'Event',style)
feuil2.write(0,2,'nA+',style)
feuil2.write(0,10,'exponent of It0 for A',style)
feuil2.write(0,0,'nA',style)
feuil1.write(0,0,'nX',style)

feuil1.write(0,11,"exponent I of RusX",style)
feuil2.write(0,11,"exponent I of RusA",style)
feuil1.write(0,3,"Test AX (onset)",style)
feuil2.write(0,7,"eA",style)
feuil1.write(0,7,"eX",style)


L1=[]
L2=[]
L3=[]
L4=[]
L=[L1,L2,L3,L4]
P1=[]
P1copie=[]
eA=1.0
eX=1.0

XIt0=20.0
AIt0=20.0
IM=XIt0
IMA=AIt0


trialX=0
trialA=0



nA=0
nAplus=0
nX=0
nXplus=0


n=0

P1US=0
P2US=0
P3US=0
firstX=0
firstA=0

trialA=0
trialX=0
trialX1=0
trialA1=0
trialA2=0
trialX2=0
trialX3=0
trialA3=0

outcome=1
MX=1.0
MA=1.0
a=0
b=0
c=0
abc=[P1US,P2US,P3US]
ABC=[a,b,c]

abc=[float(P1US),float(P2US),float(P3US)]

P1copie=[]
P1=[]
P2=[]
P2copie=[]
P3=[]
P3copie=[]
Pcopie=[]
P1US=float(input("Phase1: value of Rus between 0.1 and 1:" ))
PUS=[]

P1Xplus=input("Phase1:  number of X+:" )
for i in range(int(P1Xplus)):
    P1.append("X+")

P1Xmoins=input("Phase1:  number of X- : ")
for i in range( int(P1Xmoins)):
    P1.append("X-")

P1Aplus=input("Phase1:  number of A+ : ")
for i in range( int(P1Aplus)):
    P1.append("A+" )

P1Amoins=input("Phase1:  number of A- : ")
for i in range(int(P1Amoins)):
    P1.append("A-")

  
P2US=float(input("Phase2: value of Rus between 0.1 and 1: " ))
if P2US==P1US:
    #Pcopie=[P1copie,P2copie,P3copie]
        
    P2Xplus=input("Phase2:  number of X+ : ")
    for i in range(int(P2Xplus)):
        P2.append("X+")
   

    P2Xmoins=input("Phase2: number of X- : ")
    for i in range(int(P2Xmoins)):
        P2.append("X-")
    #if P1Xplus+P2Xplus==0:
        #XIt0=XIt0+P1Xmoins+P2Xmoins


    P2Aplus=input("Phase2:  number of A+ : ")
    for i in range( int(P2Aplus)):
        P2.append("A+")

    P2Amoins=input("Phase2:  number of A- : ")
    for i in range( int(P2Amoins)):
        P2.append("A-")
    ##if P1Aplus+P2Aplus==0:
        #AIt0=AIt0+P1Amoins+P2Amoins


    random.shuffle(P2)
    P3US=float(input("Phase3: value of Rus betweenA 0.1 and 1:  "))
    


    P3Xplus=input("Phase3:  number of X+ : ")
    for i in range( int(P3Xplus)):
        P3.append("X+")


    P3Xmoins=input("Phase3:  number of X- : ")
    for i in range( int(P3Xmoins)):
        P3.append("X-" )

    P3Aplus=input("Phase3:  number of A+ : ")
    for i in range( int(P3Aplus)):
        P3.append("A+" )

    P3Amoins=input("Phase3:  number of A- : ")
    for i in range( int(P3Amoins)):
        P3.append("A-" )

    random.shuffle(P3)

a= int(P1Xplus)+ int(P1Xmoins)+ int(P1Aplus)+ int(P1Amoins)
for i in range(a):
    P1copie.append( int(i))
if P2US!=P1US:
    rg=1
    ABC=[a]
    
    outcome=P1US
    P=[P1]
    
    random.shuffle(P1copie)
    orderPass=[P1copie]
else:
    rg=3
    
    
    P=[P1,P2,P3]
    abc=[P1US,P2US,P3US]
    b= int(P2Xplus)+ int(P2Xmoins)+ int(P2Aplus)+ int(P2Amoins)
    for i in range(b):
        P2copie.append( int(i))

    c= int(P3Xplus)+ int(P3Xmoins)+ int(P3Aplus)+ int(P3Amoins)
    for i in range(c):
        P3copie.append( int(i))

    random.shuffle(P1copie)
    random.shuffle(P2copie)
    random.shuffle(P3copie)
    
    orderPass=[P1copie,P2copie,P3copie]
    
    ABC=[a,b,c]
for k in range(len(P1copie)):
    PUS.append (P1US)
for k in range(len(P2copie)):
    PUS.append (P2US)
for k in range(len(P3copie)):
    PUS.append (P3US)

   
    
for j in range(rg):

    for i in range(ABC[j]):
        rus=PUS[n]
        if P[j][orderPass[j][i]]=="A+":
            n=n+1
            nA=nA+1
            trialA=trialA+1
            MA=float(eA/(nAplus+1))
            #MX=float(eX/(nXplus+1))
            
            
            IM=float(XIt0**MX)#MtestX)
            Lastevent="A+"
            
            
            IMA=float(AIt0**MA)
           
            IMAX=float(IMA+IM)
            CRX=float(rus**IM)
            CRA=float(rus**IMA)
            CRAX=float(rus**IMAX)
        
            
            feuil2.write(nA,1,n,style)#horiz center  ))
            feuil2.write(nA,4,P[j][orderPass[j][i]],style)
            
            feuil2.write(nA,8,rus,style)
            
            feuil2.write(nA,7, round(float(eA),6 ),style)
            feuil2.write(nA,2, round(float(nAplus),2 ),style)
            feuil2.write(nA,10, round(float(MA),6 ),style)
            feuil2.write(nA,0, round(float(nA),2) ,style)

            feuil2.write(nA,11, round(float(IMA),2 ),style)
            feuil2.write(nA,3, round(float(CRA),6 ),style4)
        
           


            if nAplus==0:
            
                firstA=float(rus** int(AIt0))
        
           
            nAplus=nAplus+1
            eA=1.0
                   
                
            
            
        elif P[j][orderPass[j][i]]=="A-":
            
            n=n+1
            nA=nA+1
            trialA=trialA+1            
            MA=float(eA/(nAplus+1))
            MX=float(eX/(nXplus+1))
            
            
            IM=float(XIt0**MX)#MtestX)
            Lastevent="A-"
            
            
            IMA=float(AIt0**MA)
           
            IMAX=float(IMA+IM)
            CRX=float(rus**IM)
            CRA=float(rus**IMA)
            CRAX=float(rus**IMAX)
            
            feuil2.write(nA,0,nA,style)#horiz center  ))
            feuil2.write(nA,4,P[j][orderPass[j][i]],style)
            
            feuil2.write(nA,8,rus,style)
            
            feuil2.write(nA,7, round(float(eA),6) ,style)
            feuil2.write(nA,2, round(float(nAplus),2) ,style)
            feuil2.write(nA,10, round(float(MA),6) ,style)
            feuil2.write(nA,1, round(float(n),2) ,style)

            feuil2.write(nA,11, round(float(IMA),2 ),style)
            feuil2.write(nA,3, round(float(CRA),6 ),style4)
            #feuil2.write(nA+1,8,rus)
        




            eA=eA+(nAplus/(nA+1.0))
        

            if nAplus==0:
            
            
                firstA=float(rus** int(AIt0))
                AIt0=AIt0+1
    
           
        
        elif P[j][orderPass[j][i]]=="X+":
            nX=nX+1
            n=n+1
            trialX=trialX+1
            MA=float(eA/(nAplus+1))
            MX=float(eX/(nXplus+1))
            
            
            IM=float(XIt0**MX)#MtestX)
            
            
            Lastevent="X+"
            IMA=float(AIt0**MA)
           
            IMAX=float(IMA+IM)
            CRX=float(rus**IM)
            CRA=float(rus**IMA)
            CRAX=float(rus**IMAX)
            

            feuil1.write(nX,4,P[j][orderPass[j][i]],style)

            feuil1.write(nX,1,n,style)
            feuil1.write(nX,7, round(float(eX),6 ),style)
            feuil1.write(nX,5, round(float(nXplus+1),2 ),style)
            feuil1.write(nX,10, round(float(MX),6 ),style)
            feuil1.write(nX,0, round(float(nX),2 ),style)
            #feuil1.write(nX+1,11,"X tested )
                    
            feuil1.write(nX,11, round(float(IM),2 ),style)
            feuil1.write(nX,2, round(float(CRX),6),style4)
            
            feuil1.write(nX,3, round(float(CRAX),6),style4 )
           
                
            feuil1.write(nX,8,float(rus ),style)
                 
            firstX=float(rus** int(XIt0))
            nXplus=nXplus+1.0
                       

            eX=1.0
           
                    
        elif P[j][orderPass[j][i]]=="X-":
            nX=nX+1
            n=n+1
            trialX=trialX+1
            MA=float(eA/(nAplus+1))
            MX=float(eX/(nXplus+1))
            
            
            IM=float(XIt0**MX)#MtestX)
            
            
            Lastevent="X-"
            IMA=float(AIt0**MA)
           
            IMAX=float(IMA+IM)
            CRX=float(rus**IM)
            CRA=float(rus**IMA)
            CRAX=float(rus**IMAX)
            

            feuil1.write(nX,4,P[j][orderPass[j][i]],style)

            feuil1.write(nX,1,n,style)
            feuil1.write(nX,7, round(float(eX),6 ),style)
            feuil1.write(nX,5, round(float(nXplus),2 ),style)
            feuil1.write(nX,10, round(float(MX),6 ),style)
            feuil1.write(nX,0, round(float(nX),2 ),style)
            #feuil1.write(nX+1,11,"X tested )
                    
            feuil1.write(nX,11, round(float(IM),2 ),style)
            feuil1.write(nX,2, round(float(CRX),6),style4)
            
            feuil1.write(nX,3, round(float(CRAX),6 ),style4)
           
                
            feuil1.write(nX,8,float(rus ),style)
                 
            eX=eX+(nXplus/(nX+1))
            if nXplus==0:
                
                firstX=float(rus** int(XIt0))
                XIt0=XIt0+1


    

if P2US!=P1US:
    topX_row= int(nX+1)
    bottomX_row= int(nX+3)
    leftX_column=4
    rightX_column=7
    topA_row= int(nA+1)
    bottomA_row= int(nA+3)
    leftA_column=4
    rightA_column=7
    #feuil1.write(nA+1,8,P2US)
    #feuil2.write(nA+1,8,P2US)
    feuil1.write_merge(topX_row,bottomX_row,leftX_column,rightX_column,"outcomes revaluation (US revaluation and SPC)",stylebis)
    feuil2.write_merge(topA_row,bottomA_row,leftA_column,rightA_column,"outcomes revaluation (US revaluation and SPC)",stylebis)

    #feuil1.write(nA+2,1,nA+1)
    CRXdev=P2US**IM
    CRAdev=P2US**IMA
    CRAXdev=P2US**(IMAX)
    feuil1.write(nX+2,2,CRXdev,style4)
    feuil2.write(nA+2,3,CRAdev,style4)

    feuil1.write(nX+2,3,CRAXdev,style4)
    H=nX+1
    B=nX+1
    HA=nA+1
    BA=nA+1
    L=0
    LL=9
    RR=11
    R=3

    feuil1.write_merge(H,B,L,R,"",style5)
    feuil2.write_merge(HA,BA,L,R,"",style5)
    feuil1.write_merge(H,B,LL,RR,"",style5)
    feuil2.write_merge(HA,BA,LL,RR,"",style5)



    feuil1.write(nX+1,8,float(P2US ),style)
    feuil2.write(nA+1,8,float(P2US ),style)

        
    feuil1.write(nX+2,1,"following US revaluation",styleter)
    feuil1.write(nX+2,0,"response to X",styleter)
    feuil2.write(nA+2,1,"following US revaluation",styleter)
    feuil2.write(nA+2,0,"response to A",styleter)

else:
    HB=nX+5
    BB=nX+5
    LLL=0
    RRR=12
##    feuil1.write_merge(HB,BB,LLL,RRR,"Note1: At the time when the first X is introduced (i.e., nX=1), nxplus is 0 because the test of X is before any CS-US pairing",styleter)
##    feuil1.write_merge(nX+7,nX+7,LLL,RRR,"Note2: At the time of the nth CS-US pairing (iE, nX=n), nX+=n-1 because onset of X precedes the US) ",styleter)
##    
##    feuil2.write_merge(HB,BB,LLL,RRR,"Note1: At the time when the first X is introduced (i.e., nX=1), nxplus is 0 because the test of X is before any CS-US pairing",styleter)
##    feuil2.write_merge(nX+7,nX+7,LLL,RRR,"Note2: At the time of the nth CS-US pairing (iE, nX=n), nX+=n-1 because onset of X precedes the US) ",styleter)
##    
  

    MA=float(eA/(nAplus+1))

    
    IMA=float(AIt0**MA)
   
    IMAX=float(IMA+IM)
    
    CRA=float(rus**IMA)

    
    
    feuil2.write(nA+1,3, round(float(CRA),6 ),style4)

    feuil2.write(nA+1,2,"A final test = ",style4)
        


    MX=float(eX/(nXplus+1))
        
        
    IM=float(XIt0**MX)#MtestX)
    
   
    IMAX=float(IMA+IM)
    CRX=float(rus**IM)

    CRAX=float(rus**IMAX)
          
    

    
    
    feuil1.write(nX+1,1,"X final test:",style4)
            
    feuil1.write(nX+1,2, round(float(CRX),6),style4)
    
    feuil1.write(nX+1,3, round(float(CRAX),6),style4 )
   
                    
   

         
print ("PLEASE WAIT: EDITION OF AN EXCEL FILE")



test.save("sim.xls")

#os.startfile("sim.xls")
os.system("open sim.xls")
