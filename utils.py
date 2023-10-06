
from pyautocad import Autocad, APoint
acad = Autocad(create_if_not_exists=True)
from math import *
import array
import win32com.client
import pythoncom
import numpy
import numpy as np
import pandas as pd

acad = win32com.client.Dispatch("AutoCAD.Application")
acadModel = acad.ActiveDocument.ModelSpace

def APoint1(x, y, z = 0):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def ADouble(xyz):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def variants(object):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (object))



def Keyway_Coordinates(SP,Side,Height,Width):
    Coordinates = []
    if Side == "Left":
            x = SP.x; y = SP.y+Height/2   
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
            x = SP.x+Width; y= SP.y+Height/2 
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
            x = SP.x+Width; y= SP.y-Height/2 
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
            x = SP.x; y= SP.y-Height/2 
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
    else:
            x = SP.x; y = SP.y+Height/2   
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
            x = SP.x-Width; y= SP.y+Height/2 
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
            x = SP.x-Width; y= SP.y-Height/2 
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
            x = SP.x; y= SP.y-Height/2 
            pnt1 = [x,y,0]
            Coordinates.extend(pnt1) 
    return Coordinates

################################################################################################################



class Keysection:
    # Constructor method (init method) to initialize attributes
    def __init__(self, SP, Side, Type, Dia, Keywidth, Keydepth):
        self.SP = SP
        self.Side = Side # Left or Right
        self.Type = Type # Open or Close
        self.Dia = Dia
        self.Keywidth = Keywidth
        self.Keydepth = Keydepth
       
    # Method to draw Keysection
    def Draw_Keysection(self): 
        out_loop = []
        in_loop = []
        
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acadModel = acad.ActiveDocument.ModelSpace
        
        def APoint1(x, y, z = 0):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

        def ADouble(xyz):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

        def variants(object):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (object))

        p1 = APoint(0,0)        
        delta = self.Dia/2 - sqrt(pow(self.Dia/2,2)-pow(self.Keywidth/2,2))
        if self.Side == "Left":
            p1.x = self.SP.x - self.Dia/2 + delta; p1.y = self.SP.y
            Coordinates = Keyway_Coordinates(p1,self.Side,self.Keywidth,self.Keydepth-delta)
            Pline = acadModel.AddPolyline(ADouble(Coordinates))
            Start_Angle = pi + asin(self.Keywidth / self.Dia)
            End_Angle = -Start_Angle 
             
        else:
            p1.x = self.SP.x + self.Dia/2 - delta; p1.y = self.SP.y
            Coordinates = Keyway_Coordinates(p1,self.Side,self.Keywidth,self.Keydepth-delta)
            Pline = acadModel.AddPolyline(ADouble(Coordinates))
            Start_Angle = asin(self.Keywidth / self.Dia)
            End_Angle = -Start_Angle 
         
        Arc1 = acadModel.AddArc(APoint1(self.SP.x, self.SP.y, 0), self.Dia/2, Start_Angle, End_Angle)
        if self.Type == "Close":
            acadModel.AddArc(APoint1(self.SP.x, self.SP.y, 0), self.Dia/2, End_Angle, Start_Angle)  
         
         ###### Hatch  ######          
       
        out_loop.append(Pline)
        out_loop.append(Arc1)
        outer = variants(out_loop)

        hatch = acadModel.AddHatch(0, "ANSI31", True,)
        hatch.AppendOuterLoop(outer)
        hatch.PatternScale = 40   
        hatch.Color = 2     
        # in_loop.append(acadModel.AddCircle(APoint1(250, 250, 0), 100))
        # inner = variants(in_loop)
        # hatch.AppendInnerLoop(inner)

        ###### Centre line  ###### 
        
        Draw_CentreLine_Plus(self.SP,self.Dia)
        
        

################################################################################################################

################################################################################################################

class Keyway:
    # Constructor method (init method) to initialize attributes
    def __init__(self, SP, Left_Shape, Right_Shape, Width, Length):
        self.SP = SP
        self.Left_Shape = Left_Shape
        self.Right_Shape = Right_Shape
        self.Width = Width
        self.Length = Length
    
    # Method to draw Keyway
    def Draw_Keyway(self):
          p1 = APoint(0,0)
          
          if self.Left_Shape == "Square" :
                Draw_Horizontal_Edge(self.SP,"Linear",self.Width,self.Width,self.Width/2,0,7,"0","Top")
                Draw_Horizontal_Edge(self.SP,"Linear",self.Width,self.Width,self.Width/2,0,7,"0","Bottom")
                Draw_Vertical_Edge(self.SP,self.Width, 0, "Top","Solid")
                Draw_Vertical_Edge(self.SP,self.Width, 0, "Bottom","Solid")
          elif self.Left_Shape == "Round" :  
                p1.x = self.SP.x + self.Width/2; p1.y = self.SP.y
                acad.model.AddArc(p1, self.Width/2, 90*pi/180 , 270*pi/180)          
          if self.Right_Shape == "Square" :
                p1.x = self.SP.x + self.Length-self.Width/2; p1.y = self.SP.y
                Draw_Horizontal_Edge(p1,"Linear",self.Width,self.Width,self.Width/2,0,7,"0","Top")
                Draw_Horizontal_Edge(p1,"Linear",self.Width,self.Width,self.Width/2,0,7,"0","Bottom")
                p1.x = self.SP.x + self.Length; p1.y = self.SP.y
                Draw_Vertical_Edge(p1,self.Width, 0,"Top","Solid")               
                Draw_Vertical_Edge(p1,self.Width, 0,"Bottom","Solid")
          elif self.Right_Shape == "Round" :  
                p1.x = self.SP.x + self.Length - self.Width/2; p1.y = self.SP.y
                acad.model.AddArc(p1, self.Width/2, 270*pi/180 , 90*pi/180)

          p1.x = self.SP.x+self.Width/2; p1.y = self.SP.y
          Draw_Horizontal_Edge(p1,"Linear",self.Width,self.Width,self.Length-self.Width,0,7,"0","Top")
          Draw_Horizontal_Edge(p1,"Linear",self.Width,self.Width,self.Length-self.Width,0,7,"0","Bottom")

################################################################################################################

################################################################################################################

def Draw_Vertical_Edge(SP,Dia1,Dia2,Key_Width,Viewtype,Int_Ext):   # SP start point, Viewtype = Nosection, Semisection, Fullsection, Int_Ext = Internal or External edge?
                p1 = APoint(0,0)
                p2 = APoint(0,0)
                p3 = APoint(0,0)
                p4 = APoint(0,0)
                MajorDia = max(Dia1,Dia2)
                MinorDia = min(Dia1,Dia2)
                Coordinates = []

                p1.x = SP.x; p1.y = SP.y + MajorDia/2
                p3.x = SP.x; p3.y = SP.y - MajorDia/2

                if Int_Ext == "External":
                    #   TopLineType = "" 
                    #   BottomLineType = ""
                      
                      if Viewtype == "Nosection":
                            # p1.x = SP.x; p1.y = SP.y + MajorDia/2
                            p2.x = SP.x; p2.y = SP.y + Key_Width/2                            
                            # p3.x = SP.x; p3.y = SP.y - MajorDia/2
                            p4.x = SP.x; p2.y = SP.y - Key_Width/2                            
                      elif Viewtype == "Fullsection":
                            # p1.x = SP.x; p1.y = SP.y + MajorDia/2
                            p2.x = SP.x; p2.y = SP.y + MinorDia/2
                            # p3.x = SP.x; p3.y = SP.y - MajorDia/2
                            p4.x = SP.x; p4.y = SP.y - MinorDia/2
                      else:      
                            # p1.x = SP.x; p1.y = SP.y + MajorDia/2
                            p2.x = SP.x; p2.y = SP.y + MinorDia/2
                            # p3.x = SP.x; p3.y = SP.y - MajorDia/2
                            p4.x = SP.x; p2.y = SP.y - Key_Width/2    
                else:
                      
                      if Viewtype == "Nosection":
                            # TopLineType = "H2" 
                            # BottomLineType = "H2"
                            # p1.x = SP.x; p1.y = SP.y + MajorDia/2
                            p2 = SP                            
                            # p3.x = SP.x; p3.y = SP.y - MajorDia/2
                            p4 = SP                            
                      elif Viewtype == "Fullsection":
                            # TopLineType = "" 
                            # BottomLineType = ""
                            # p1.x = SP.x; p1.y = SP.y + MajorDia/2
                            p2.x = SP.x; p2.y = SP.y + MinorDia/2
                            # p3.x = SP.x; p3.y = SP.y - MajorDia/2
                            p4.x = SP.x; p4.y = SP.y - MinorDia/2

                            Coordinates.extend([p2.x,p2.y,0]) 
                            Coordinates.extend([SP.x,SP.y,0]) 
                            L3 = DrawLine(Coordinates,"")  # Extra line
                            Coordinates.clear()

                            Coordinates.extend([p4.x,p4.y,0]) 
                            Coordinates.extend([SP.x,SP.y,0]) 
                            L4 = DrawLine(Coordinates,"")  # Extra line
                            Coordinates.clear()
                            # L3 = DrawLine(p2,SP,"") # Extra line
                            # L4 = DrawLine(p4,SP,"") # Extra line
                      else:      
                            # TopLineType = "" 
                            # BottomLineType = "H2"
                            # p1.x = SP.x; p1.y = SP.y + MajorDia/2
                            p2.x = SP.x; p2.y = SP.y + MinorDia/2
                            # p3.x = SP.x; p3.y = SP.y - MajorDia/2
                            p4 = SP

                            Coordinates.extend([p2.x,p2.y,0]) 
                            Coordinates.extend([SP.x,SP.y,0]) 
                            L3 = DrawLine(Coordinates,"")  # Extra line
                            Coordinates.clear()

                            # L3 = DrawLine(p2,SP,"") # Extra line
                TopLineType,BottomLineType =  LineType(Viewtype,Int_Ext)

                Coordinates.extend([p1.x,p1.y,0]) 
                Coordinates.extend([p2.x,p2.y,0]) 
                L1 = DrawLine(Coordinates,TopLineType)  # Top line  
                Coordinates.clear()

                Coordinates.extend([p3.x,p3.y,0]) 
                Coordinates.extend([p4.x,p4.y,0]) 
                L2 = DrawLine(Coordinates,BottomLineType)  # Bottom line  
                Coordinates.clear()
                # L1 = DrawLine(p1,p2,TopLineType)   # Top line  
                # L2 = DrawLine(p3,p4,BottomLineType)   # Bottom line  
                               
                return(L1,L2)
                      


################################################################################################################
def DrawLine(Coordinates,Type):  #Type = H2 - Hidden, L2 - Centre line, 2 - yellow solid, "" - white solid
      
        
      L1 = acadModel.AddPolyline(ADouble(Coordinates))
      if Type == "H2": L1.color = 2; L1.linetype = "H2"
      elif Type == "L2": L1.color = 2; L1.linetype = "C2" 
      elif Type == "2": L1.color = 2;     
      L1.Linetypescale = 1 
      return(L1)
################################################################################################################     
def DrawArc(Coordinates,Radius,Start_Angle, End_Angle,Type):   #Type = H2 - Hidden, L2 - Centre line, 2 - yellow solid, "" - white solid
      SP = APoint(0,0) 
      SP.x = Coordinates[0]; SP.y = Coordinates[1]
      A1 = acadModel.AddArc(APoint1(SP.x, SP.y, 0), Radius, Start_Angle*pi/180, End_Angle*pi/180) 
      if Type == "H2": A1.color = 2; A1.linetype = "H2"
      elif Type == "L2": A1.color = 2; A1.linetype = "C2"  
      elif Type == "2": A1.color = 2;     
      A1.Linetypescale = 1       
      return(A1)
################################################################################################################


class StepEnd:
    # Constructor method (init method) to initialize attributes
    def __init__(self, SP, Side, Dia, Chamfer_Size, Chamfer_Angle, Break_Size,Linetype,Position):
        self.SP = SP
        self.Side = Side # Left or Right
        self.Dia = Dia
        self.Chamfer_Size = Chamfer_Size # x length
        self.Chamfer_Angle = Chamfer_Angle
        self.Break_Size = Break_Size
        self.Linetype = Linetype
        self.Position = Position
            
    # Method to draw StepEnd
    def Draw_StepEnd(self): 
        p1 = APoint(0,0)
        
        Chamfer_x = self.Chamfer_Size
        Chamfer_y = tan(self.Chamfer_Angle*pi/180) * self.Chamfer_Size
        if self.Linetype == "H2": Colorindex = 2
        else: Colorindex = 7
              
        if self.Side == "Left" :
                
                # Draw_Vertical_Edge(self.SP,self.Dia-2*Chamfer_y, self.Break_Size)
                p1.x = self.SP.x + Chamfer_x; p1.y = self.SP.y
                # Draw_Vertical_Edge(p1, self.Dia, self.Break_Size)

                Draw_Horizontal_Edge(self.SP,"Linear",self.Dia-2*Chamfer_y,self.Dia,Chamfer_x,0,Colorindex,self.Linetype,self.Position)                
                
        else :
                # Draw_Vertical_Edge(self.SP,self.Dia, self.Break_Size)
                p1.x = self.SP.x + Chamfer_x; p1.y = self.SP.y
                # Draw_Vertical_Edge(p1, self.Dia-2*Chamfer_y, self.Break_Size)

                Draw_Horizontal_Edge(self.SP,"Linear",self.Dia,self.Dia-2*Chamfer_y,Chamfer_x,0,Colorindex,self.Linetype,self.Position) 
                       

################################################################################################################


class Groove:
    # Constructor method (init method) to initialize attributes
    def __init__(self, SP, Type, Width, Depth, Step_Dia, Taper_Angle, Key_Width, ViewType, Int_Ext): #SP is the center of the groove; Type = Circular or Square, (for circular radius = width /2 and ignore depth)
        self.SP = SP
        self.Type = Type # Circular or Square
        self.Width = Width
        self.Depth = Depth
        self.Step_Dia = Step_Dia
        self.Taper_Angle = Taper_Angle
        self.Key_Width = Key_Width
        self.ViewType = ViewType
        self.Int_Ext = Int_Ext
        
    
    # Method to draw Groove
    def Draw_Groove(self): 
        # p1,p2,p3,p4,p5,p6 = APoint(0,0,0)
        p1 = APoint(0,0)
        p2 = APoint(0,0)
        # p3 = APoint(0,0,0)
        # p4 = APoint(0,0,0)
        # p5 = APoint(0,0,0)
        # p6 = APoint(0,0,0)

        delta = self.Width/2 * tan(self.Taper_Angle*pi/180)
        if self.Int_Ext == "External": Flag = 1; Angle1 = 180; Angle2 = 0
        else: Flag = -1; Angle1 = 0; Angle2 = 180

        
        
        Coordinates = []
        Coordinates1 = []
        Coordinates2 = []
        Coordinates3 = []
        
        lineTypeTop, lineTypeBottom = LineType(self.ViewType,self.Int_Ext)

        if self.Type == "Square" :
                
                Groove_Dia = self.Step_Dia - Flag * 2 * self.Depth
                
                Coordinates.extend([self.SP.x - self.Width/2,self.SP.y + self.Step_Dia/2-delta,0]) 
                Coordinates.extend([self.SP.x - self.Width/2,self.SP.y + Groove_Dia/2,0]) 
                Coordinates.extend([self.SP.x + self.Width/2,self.SP.y + Groove_Dia/2,0]) 
                Coordinates.extend([self.SP.x + self.Width/2,self.SP.y + self.Step_Dia/2+delta,0]) 
                L1 = DrawLine(Coordinates,lineTypeTop)
                Coordinates.clear()

                Coordinates.extend([self.SP.x - self.Width/2,self.SP.y - self.Step_Dia/2+delta,0]) 
                Coordinates.extend([self.SP.x - self.Width/2,self.SP.y - Groove_Dia/2,0]) 
                Coordinates.extend([self.SP.x + self.Width/2,self.SP.y - Groove_Dia/2,0]) 
                Coordinates.extend([self.SP.x + self.Width/2,self.SP.y - self.Step_Dia/2-delta,0]) 
                L2 = DrawLine(Coordinates,lineTypeBottom)
                Coordinates.clear()

                

                # Draw_Vertical_Edge(self.SP,self.Step_Dia, self.Break_Size)
                # p1.x = self.SP.x + self.Width; p1.y = self.SP.y
                # Draw_Vertical_Edge(p1, self.Step_Dia, self.Break_Size)

                # Draw_Horizontal_Edge(self.SP,"Linear",self.Groove_Dia,self.Groove_Dia,self.Width,0,7)
                
        else:
                Groove_Dia = self.Step_Dia 
                p1.x = self.SP.x; p1.y = self.SP.y + self.Step_Dia/2
                L1 = DrawArc(p1,self.Width/2,Angle1 + self.Taper_Angle, Angle2 + self.Taper_Angle,lineTypeTop)

                p2.x = self.SP.x; p2.y = self.SP.y - self.Step_Dia/2
                L2 = DrawArc(p2,self.Width/2,Angle2 - self.Taper_Angle, Angle1 - self.Taper_Angle,lineTypeBottom)

                # Draw_Vertical_Edge(self.SP,self.Step_Dia, self.Break_Size)
                # p1.x = self.SP.x + self.Width; p1.y = self.SP.y
                # Draw_Vertical_Edge(p1, self.Step_Dia, self.Break_Size)
                
                # Draw_Horizontal_Edge(self.SP,"Arc",self.Groove_Dia+self.Width, self.Width/2,180,0,7)       
        
# ExtraLine
        p1.x = self.SP.x - self.Width/2; p1.y = self.SP.y
        p2.x = self.SP.x + self.Width/2; p2.y = self.SP.y
        Vertical_Edge_Line(p1, Groove_Dia, self.Key_Width, self.ViewType, self.Int_Ext)
        Vertical_Edge_Line(p2, Groove_Dia, self.Key_Width, self.ViewType, self.Int_Ext)

      #   Coordinates.extend([self.SP.x - self.Width/2,self.SP.y + Groove_Dia/2,0])
      #   Coordinates.extend([self.SP.x - self.Width/2,self.SP.y + self.Key_Width/2,0])

      #   Coordinates1.extend([self.SP.x + self.Width/2,self.SP.y + Groove_Dia/2,0])
      #   Coordinates1.extend([self.SP.x + self.Width/2,self.SP.y + self.Key_Width/2,0])

      #   Coordinates2.extend([self.SP.x - self.Width/2,self.SP.y - Groove_Dia/2,0])
      #   Coordinates2.extend([self.SP.x - self.Width/2,self.SP.y - self.Key_Width/2,0])

      #   Coordinates3.extend([self.SP.x + self.Width/2,self.SP.y - Groove_Dia/2,0])
      #   Coordinates3.extend([self.SP.x + self.Width/2,self.SP.y - self.Key_Width/2,0])

      #   if self.Int_Ext == "External":
      #           if self.ViewType == "Nosection":
      #                       DrawLine(Coordinates,"")
      #                       DrawLine(Coordinates1,"")
      #                       DrawLine(Coordinates2,"")
      #                       DrawLine(Coordinates3,"")                                                  
      #           elif self.ViewType == "Semisection":                            
      #                       DrawLine(Coordinates2,"")
      #                       DrawLine(Coordinates3,"")
                     
      #   else:      
      #           if self.ViewType == "Nosection":
      #                       DrawLine(Coordinates,"H2")
      #                       DrawLine(Coordinates1,"H2")
      #                       DrawLine(Coordinates2,"H2")
      #                       DrawLine(Coordinates3,"H2")                                                  
      #           elif self.ViewType == "Semisection":
      #                       DrawLine(Coordinates,"")
      #                       DrawLine(Coordinates1,"")                            
      #                       DrawLine(Coordinates2,"H2")
      #                       DrawLine(Coordinates3,"H2")
      #           else:
      #                       DrawLine(Coordinates,"")
      #                       DrawLine(Coordinates1,"")                            
      #                       DrawLine(Coordinates2,"")
      #                       DrawLine(Coordinates3,"")
                
      #   Coordinates.clear()
      #   Coordinates1.clear()
      #   Coordinates2.clear()
      #   Coordinates3.clear()
    # End of ExtraLine                         
        return (L1,L2)        

################################################################################################################


class Radius:
    # Constructor method (init method) to initialize attributes
    def __init__(self, SP, Dia1, Dia2, Radius, Side ,TaperAngle, KeyWidth, Viewtype,Int_Ext):
        self.SP = SP
        self.Dia1 = Dia1
        self.Dia2 = Dia2
        self.Radius = Radius        
        self.Side = Side # Left or Right
        self.TaperAngle = TaperAngle
        self.KeyWidth = KeyWidth
        self.Viewtype = Viewtype
        self.Int_Ext = Int_Ext
       
    # Method to draw Radius
    def Draw_Radius(self): 
        p1 = APoint(0,0)
        p2 = APoint(0,0)

        

        delta = delta1(self.Radius,self.TaperAngle) - delta2(self.Radius,self.TaperAngle)
        
        lineTypeTop, lineTypeBottom = LineType(self.Viewtype,self.Int_Ext)
        if self.Dia1>self.Dia2:
               if self.Side == "Left":
                    Angle1 = 180; Angle2 = 270 + self.TaperAngle
                    Angle3 = 90 - self.TaperAngle; Angle4 = 180
                    p1.x = p2.x = self.SP.x + self.Radius; 
                    p1.y = self.SP.y + (self.Dia2/2 + delta); p2.y = self.SP.y - (self.Dia2/2 + delta)
               else:
                    Angle1 = 0; Angle2 = 90 + self.TaperAngle
                    Angle3 = 270 - self.TaperAngle; Angle4 = 0
                    p1.x = p2.x = self.SP.x - self.Radius; 
                    p1.y = self.SP.y - (self.Dia1/2 + delta); p2.y = self.SP.y + (self.Dia1/2 + delta)
        else:
              if self.Side == "Left":
                    Angle1 = 90 + self.TaperAngle; Angle2 = 180
                    Angle3 = 180; Angle4 = 270 - self.TaperAngle
                    p1.x = p2.x = self.SP.x + self.Radius; 
                    p1.y = self.SP.y - (self.Dia2/2 + delta); p2.y = self.SP.y + (self.Dia2/2 + delta)
              else:
                    Angle1 = 270 + self.TaperAngle; Angle2 = 0
                    Angle3 = 0; Angle4 = 90 - self.TaperAngle
                    p1.x = p2.x = self.SP.x - self.Radius; 
                    p1.y = self.SP.y + (self.Dia2/2 + delta); p2.y = self.SP.y - (self.Dia2/2 + delta)

        R1 = DrawArc(p1,self.Radius,Angle1, Angle2,lineTypeTop)    
        R2 = DrawArc(p2,self.Radius,Angle3, Angle4,lineTypeBottom)  

      #    # ExtraLine
      #   p1.x = self.SP.x - self.Chamfersize; p1.y = self.SP.y      
      #   Vertical_Edge_Line(p1, Dia, self.KeyWidth, self.Viewtype, self.Int_Ext)

        return (R1,R2)
        # if self.Linetype == "H2": Colorindex = 2
        # else: Colorindex = 7

        # if self.Side == "Left":
        #     p1.x = self.SP.x; p1.y = self.SP.y + self.Dia/2 + self.Size 
        #     Draw_Horizontal_Edge(p1,"Circular",60,4,180,0,7,self.Linetype,self.Position)
        #     p1.x = self.SP.x; p1.y = self.SP.y + self.Dia/2 + self.Size
        #     acad.model.AddArc(p1, self.Size, 270*pi/180 , 0*pi/180)  
        #     p1.x = self.SP.x; p1.y = self.SP.y - self.Dia/2 - self.Size
        #     acad.model.AddArc(p1, self.Size, 0*pi/180 , 90*pi/180)  
        # else:
        #     p1.x = self.SP.x+self.Size; p1.y = self.SP.y + self.Dia/2 + self.Size
        #     acad.model.AddArc(p1, self.Size, 180*pi/180 , 270*pi/180)  
        #     p1.x = self.SP.x+self.Size; p1.y = self.SP.y - self.Dia/2 - self.Size
        #     acad.model.AddArc(p1, self.Size, 90*pi/180 , 180*pi/180)    
              
            

################################################################################################################
def Radius_Coordinates(SP, Dia1, Dia2, Radius, Side ,TaperAngle, Position):
        Coordinates = []
        delta_1 = delta1(Radius,TaperAngle) 
        delta_2 = delta2(Radius,TaperAngle)
        Flag = getFlag(Position)
    
        if Side == "Left":
               x = SP.x + Radius
               if Dia1>Dia2:
                      y = SP.y + (Dia2/2 + (delta_1 + delta_2))*Flag
                      Start_Angle_Top = 180; Start_Angle_Bot = 90 - TaperAngle
                      End_Angle_Top = 270 + TaperAngle; End_Angle_Bot = 180
               else:
                      y = SP.y + (Dia2/2 - (delta_1 - delta_2))*Flag
                      Start_Angle_Top = 90 + TaperAngle; Start_Angle_Bot = 180
                      End_Angle_Top = 180; End_Angle_Bot = 270 - TaperAngle
        else:
               x = SP.x - Radius
               if Dia1>Dia2:
                      y = SP.y + (Dia1/2 - (delta_1 + delta_2))*Flag
                      Start_Angle_Top = 0; Start_Angle_Bot = 270 - TaperAngle
                      End_Angle_Top = 90 + TaperAngle; End_Angle_Bot = 0
               else:
                      y = SP.y + (Dia1/2 + (delta_1 - delta_2))*Flag
                      Start_Angle_Top = 270 + TaperAngle; Start_Angle_Bot = 0
                      End_Angle_Top = 0; End_Angle_Bot = 90 - TaperAngle
        
        if Position  == "Top":
              Start_Angle = Start_Angle_Top; End_Angle = End_Angle_Top
        else:
              Start_Angle = Start_Angle_Bot; End_Angle = End_Angle_Bot
        pnt = [x,y,0]
        Coordinates.extend(pnt) 
        return Coordinates, Start_Angle, End_Angle

################################################################################################################
def Chamfer_Coordinates(SP, Dia1, Dia2, Chamfer, Side ,TaperAngle, Position):
        Coordinates = []
        delta_1 = delta1(Chamfer,TaperAngle) 
        delta_2 = delta2(Chamfer,TaperAngle)
        Flag = getFlag(Position)

        if Side == "Left":
               x1 = SP.x; x2 = SP.x + Chamfer
               if Dia1>Dia2:
                      y1 = SP.y + (Dia2/2 + delta_1)*Flag
                      y2 = SP.y + (Dia2/2 + delta_2)*Flag
               else:
                      y1 = SP.y + (Dia2/2 - delta_1)*Flag
                      y2 = SP.y + (Dia2/2 + delta_2)*Flag
        else:
               x1 = SP.x - Chamfer; x2 = SP.x
               if Dia1>Dia2:
                      y1 = SP.y + (Dia1/2 - delta_2)*Flag 
                      y2 = SP.y + (Dia1/2 - delta_1)*Flag
                                           
               else:
                      y1 = SP.y + (Dia1/2 - delta_2)*Flag
                      y2 = SP.y + (Dia1/2 + delta_1)*Flag
                      
                      
        pnt1 = [x1,y1,0]
        Coordinates.extend(pnt1) 
        pnt2 = [x2,y2,0]
        Coordinates.extend(pnt2)
        return Coordinates

################################################################################################################
def End_Coordinates(SP, Dia, Position):
      Coordinates = []
   
      Flag = getFlag(Position)
      pnt = [SP.x,SP.y + (Dia/2)*Flag,0]
      Coordinates.extend(pnt)
      return Coordinates           
################################################################################################################
def getFlag(Position):
    if Position  == "Top":
              Flag  = 1
    else:
              Flag = -1  
    return Flag 
################################################################################################################
def Draw_CentreLine_Plus(SP,Dia):
                p1 = APoint(0,0)
                p2 = APoint(0,0)

                p1.x = SP.x - 1.2*Dia/2; p1.y = SP.y
                p2.x = SP.x + 1.2*Dia / 2; p2.y = SP.y
                Line1 = acad.model.AddLine(p1, p2)
                Line1.Color = 2
                Line1.Linetype = "C2"
                Line1.Linetypescale = 0.5

                p1.x = SP.x; p1.y = SP.y - 1.2*Dia/2
                p2.x = SP.x; p2.y = SP.y + 1.2*Dia / 2
                Line2 = acad.model.AddLine(p1, p2)
                Line2.Color = 2  
                Line2.Linetype = "C2"  
                Line2.Linetypescale = 0.5

################################################################################################################


def Insert_Text(SP,Text,Type):
                if Type == "Big":
                      Font = "ROMANT"
                      Color = 7
                      Height = 3.0
                      Widthfactor = 1.0
                else:
                      Font = "ROMANS"
                      Color = 3
                      Height = 2.25
                      Widthfactor = 0.75
                Text1 = acad.model.AddText(Text, SP, Height)
                Text1.StyleName = Font
                Text1.Color = Color
                # Text1.Width = Widthfactor
               
################################################################################################################


def Insert_Block(SP,Blockname,X_Scalefactor,Y_Scalefactor,Z_Scalefactor,Rotation_Angle):
       
       def POINT(x,y,z):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x,y,z))
         
       acad = win32com.client.Dispatch("AutoCAD.Application")
    #    acadModel = acad.ActiveDocument.ModelSpace
       acad.ActiveDocument.ModelSpace.InsertBlock(POINT(SP.x,SP.y,0.0), Blockname, X_Scalefactor,Y_Scalefactor,Z_Scalefactor,Rotation_Angle*pi/180)
                              
################################################################################################################

def Insert_LinDim(SP,EP,Loc,Prefix,Suffix,Tol_UL,Tol_LL,Override,Dimlfac):
      TP = APoint(0,0)
      if SP.y == EP.y: #Horizontal
            TP.x = (EP.x-SP.x)/2; TP.y = SP.y + Loc
      if SP.x == EP.x: #Vertical
            TP.x = SP.x + Loc; TP.y = (EP.y - SP.y)/2  
            
      dim1 = acad.model.AddDimAligned(SP, EP, TP)
      dim1.LinearScaleFactor = Dimlfac
      dim1.SuppressTrailingZeros = True
      dim1.TextColor = 3
      dim1.TextPrefix = Prefix
      dim1.ToleranceDisplay = int(Tol_UL != 0 or Tol_LL != 0)
      dim1.ToleranceUpperLimit = Tol_UL
      dim1.ToleranceLowerLimit = Tol_LL
      dim1.TolerancePrecision = 3
      dim1.TextSuffix = Suffix
      dim1.Textoverride = Override
      

################################################################################################################

def Insert_RadialDim(CP,Radius,Angle,Length,Override,Dimlfac):     #Center point CP, Radius, Angle 0 to 360 ACW, Length of leader line, Dimlfac
      ChordPoint = APoint(0,0)
      ChordPoint.x = CP.x + Radius*cos(pi*Angle/180); ChordPoint.y = CP.y + Radius*sin(pi*Angle/180)    
            
      dim1 = acad.model.AddDimRadial(CP, ChordPoint, Length)
      dim1.LinearScaleFactor = Dimlfac
      dim1.SuppressTrailingZeros = True
      dim1.TextColor = 3
    #   dim1.TextPrefix = Prefix
    #   dim1.ToleranceDisplay = 1
    #   dim1.ToleranceUpperLimit = 0.3
    #   dim1.ToleranceLowerLimit = 0.4
    #   dim1.TolerancePrecision = 3
    #   dim1.TextSuffix = Suffix
      dim1.Textoverride = Override
            
################################################################################################################

def Insert_GrooveDetail(SP,Type,Int_Ext,Size):
    P1 = APoint(0,0)
    P2 = APoint(0,0)

    if Type == "Oring":
        P1.x = SP.x-12; P1.y = SP.y-20
        P2.x = SP.x+12; P2.y = SP.y-20
        Insert_Block(SP,"GROOVE_DET_OR",1,1,1,0) 
        Insert_LinDim(P1,P2,-35,"","",0,0,Size,1)
        if Int_Ext == "Internal":
            Insert_Block(SP,"GROOVE_DET_OR_HATCH",1,1,1,0) 
        
    if Type == "Circular":
        Insert_Block(SP,"GROOVE_DET_CI",1,1,1,0) 
        Insert_RadialDim(SP,6,-45,25,"R"+str(Size),1)
        if Int_Ext == "Internal":
            Insert_Block(SP,"GROOVE_DET_CI_HATCH",1,1,1,0)   
    
    # p1 = APoint(0,0)
    # p2 = APoint(0,0)
    # delta = 10

    # if Type == "Circular":
    #    p1.y = p2.y = SP.y

    #    p1.x = SP.x-Size-delta
    #    p2.x = SP.x-Size
    #    acad.model.AddLine(p1, p2)
       
    #    p1.x = SP.x+Size+delta
    #    p2.x = SP.x+Size
    #    acad.model.AddLine(p1, p2)
         
    #    arc1 = acad.model.AddArc(SP, Size, 180*pi/180 , 0*pi/180)
    #    arc2 = acad.model.AddArc(SP, Size+delta, 180*pi/180 , 0*pi/180)
    #    arc2.color = 2
    
    # if Type == "Oring":
    #    if Size<5: r1,r2 = 0.5, 0.2  
    #    else: r1,r2 = 1, 0.3   
 

################################################################################################################
def LineType1(Viewtype,Int_Ext):
      
    if Int_Ext == "External":
                      TopLineType = "" 
                      BottomLineType = ""                      
    else:
                      
                      if Viewtype == "Nosection":
                            TopLineType = "H2" 
                            BottomLineType = "H2"                                                    
                      elif Viewtype == "Fullsection":
                            TopLineType = "" 
                            BottomLineType = ""
                      else:      
                            TopLineType = "" 
                            BottomLineType = "H2"
    return(TopLineType,BottomLineType)                            
################################################################################################################
def LineType(Viewtype,Int_Ext):
      if Int_Ext == "External":
                      TopLineType = "L2" 
                      BottomLineType = "L2"                      
      else:
                      
                      if Viewtype == "Nosection":
                            TopLineType = "H2" 
                            BottomLineType = "H2"                                                    
                      elif Viewtype == "Fullsection":
                            TopLineType = "L2" 
                            BottomLineType = "L2"
                      else:      
                            TopLineType = "L2" 
                            BottomLineType = "H2"

      # Extra line Flag 
      if Int_Ext == "External":
             if Viewtype == "Nosection":
                    TopFlag = 1; BottomFlag = 1
             if Viewtype == "Fullsection":
                    TopFlag = 0; BottomFlag = 0
             else:
                    TopFlag = 0; BottomFlag = 1
      else:
             TopFlag = 1; BottomFlag = 1       
      return TopLineType,BottomLineType,TopFlag,BottomFlag
################################################################################################################
class Chamfer:
    # Constructor method (init method) to initialize attributes
    def __init__(self, SP, Dia1, Dia2, ChamferSize, Side ,TaperAngle, Key_Width, Viewtype,Int_Ext):
        self.SP = SP
        self.Dia1 = Dia1
        self.Dia2 = Dia2
        self.Chamfersize = ChamferSize        
        self.Side = Side # Left or Right
        self.TaperAngle = TaperAngle
        self.Key_Width = Key_Width
        self.Viewtype = Viewtype
        self.Int_Ext = Int_Ext
       
    # Method to draw Chamfer
    def Draw_Chamfer(self): 
      p1 = APoint(0,0)
      p2 = APoint(0,0)
      Coordinates = []    
      Coordinates1 = []

      delta_1 = delta1(self.Chamfersize,self.TaperAngle)
      delta_2 = delta2(self.Chamfersize,self.TaperAngle)  
      delta = delta_1 - delta_2

      if self.Dia1 > self.Dia2: Flag = -1
      else: Flag = 1
      
        
      lineTypeTop, lineTypeBottom = LineType(self.Viewtype,self.Int_Ext)
      if self.Dia1>self.Dia2:
               if self.Side == "Left":
                    Coordinates.extend([self.SP.x,self.SP.y + self.Dia2/2 + self.Chamfersize,0]) 
                    Coordinates.extend([self.SP.x + self.Chamfersize,self.SP.y + self.Dia2/2 + delta_2,0])      
                    
                    Coordinates1.extend([self.SP.x,self.SP.y - self.Dia2/2 - self.Chamfersize,0]) 
                    Coordinates1.extend([self.SP.x + self.Chamfersize,self.SP.y - self.Dia2/2 - delta_2,0]) 

                    Dia = self.Dia2 + 2*delta_2
                   
                    
               else:
                    Coordinates.extend([self.SP.x,self.SP.y + self.Dia1/2 - self.Chamfersize,0]) 
                    Coordinates.extend([self.SP.x - self.Chamfersize,self.SP.y + self.Dia1/2 - delta_2,0])                     

                    Coordinates1.extend([self.SP.x,self.SP.y - self.Dia1/2 + self.Chamfersize,0]) 
                    Coordinates1.extend([self.SP.x - self.Chamfersize,self.SP.y - self.Dia1/2 + delta_2,0])  

                    Dia = self.Dia1 - 2*delta_2
                  
      else:
              if self.Side == "Left":
                    Coordinates.extend([self.SP.x,self.SP.y + self.Dia2/2 - self.Chamfersize,0]) 
                    Coordinates.extend([self.SP.x + self.Chamfersize,self.SP.y + self.Dia2/2 + delta_2,0])                     

                    Coordinates1.extend([self.SP.x,self.SP.y - self.Dia2/2 + self.Chamfersize,0]) 
                    Coordinates1.extend([self.SP.x + self.Chamfersize,self.SP.y - self.Dia2/2 - delta_2,0]) 

                    Dia = self.Dia2 + 2*delta_2
                   
              else:
                    Coordinates.extend([self.SP.x,self.SP.y + self.Dia1/2 + self.Chamfersize,0]) 
                    Coordinates.extend([self.SP.x - self.Chamfersize,self.SP.y + self.Dia1/2 - delta_2,0])                     

                    Coordinates1.extend([self.SP.x,self.SP.y - self.Dia1/2 - self.Chamfersize,0]) 
                    Coordinates1.extend([self.SP.x - self.Chamfersize,self.SP.y - self.Dia1/2 + delta_2,0]) 

                    Dia = self.Dia1 - 2*delta_2
                    

      L1 = DrawLine(Coordinates,lineTypeTop)
      Coordinates.clear()

      L2 = DrawLine(Coordinates1,lineTypeBottom)
      Coordinates1.clear()

      # ExtraLine
      p1.x = self.SP.x + self.Chamfersize * Flag; p1.y = self.SP.y      
      Vertical_Edge_Line(p1, Dia, self.Key_Width, self.Viewtype, self.Int_Ext)

      
      # p1.x = self.SP.x; p1.y = self.SP.y      
      # Vertical_Edge_Line(p1, Dia1, self.Key_Width, self.Viewtype, self.Int_Ext)
      
      return (L1,L2)

################################################################################################################
def delta1(len,angle):
       return(len/cos(angle*pi/180))

################################################################################################################
def delta2(len,angle):
       return(len*tan(angle*pi/180))

################################################################################################################

def Vertical_Edge_Line(SP1, Diameter, Key_Width, ViewType, Int_Ext):
       Coordinates = []
       Coordinates1 = []
      

       Coordinates.extend([SP1.x,SP1.y + Diameter/2,0])
       Coordinates.extend([SP1.x,SP1.y + Key_Width/2,0])

       Coordinates1.extend([SP1.x,SP1.y - Diameter/2,0])
       Coordinates1.extend([SP1.x,SP1.y - Key_Width/2,0])


   

       if Int_Ext == "External":
                if ViewType == "Nosection":
                            DrawLine(Coordinates,"")
                            DrawLine(Coordinates1,"")   
                elif ViewType == "Semisection":
                            
                            DrawLine(Coordinates1,"")                                                                          
                                    
       else:      
                if ViewType == "Nosection":
                            DrawLine(Coordinates,"H2")
                            DrawLine(Coordinates1,"H2")
                                                                             
                elif ViewType == "Fullsection":
                            DrawLine(Coordinates,"")
                            DrawLine(Coordinates1,"")                            
                            
                else:
                            DrawLine(Coordinates,"")
                            DrawLine(Coordinates1,"H2")                            
                            
                
       Coordinates.clear()
       Coordinates1.clear()

################################################################################################################

def Taper_Angle(Dia1, Dia2, Len):
      
       return(numpy.arctan((Dia2 - Dia1)/(2*Len))*180/pi)
################################################################################################################


def Key_Width_List(I_E,Len,Key_Start,Key_Length,Key_Width):
       absKeyPosStart_Ext = 0  
       keyPosArray_Ext = []
       absKeyPosStart_Int = 0  
       keyPosArray_Int = []
      #  Distance = numpy.float64(Distance)
       for j in range(0,len(I_E)): 
            if I_E[j] =="External":
                  if pd.isna(Key_Start[j]) is False:
                        keyPosArray_Ext.append([absKeyPosStart_Ext + Key_Start[j],absKeyPosStart_Ext + Key_Start[j] + Key_Length[j],Key_Width[j]])
                  absKeyPosStart_Ext = absKeyPosStart_Ext + Len[j]
            else:
                  if pd.isna(Key_Start[j]) is False:
                        keyPosArray_Int.append([absKeyPosStart_Int + Key_Start[j],absKeyPosStart_Int + Key_Start[j] + Key_Length[j],Key_Width[j]])
                  absKeyPosStart_Int = absKeyPosStart_Int + Len[j]     
            
       return(keyPosArray_Ext, keyPosArray_Int)

################################################################################################################


def Identify_Key_Width(keyPosArray_Ext, keyPosArray_Int, Distance):
       for i in range(0,len(keyPosArray_Ext)): 
            if keyPosArray_Ext[i][0] < Distance < keyPosArray_Ext[i][1]:
                a = keyPosArray_Ext[i][2]; break
            else:
                a = 0
            
       for i in range(0,len(keyPosArray_Int)): 
            if keyPosArray_Int[i][0] < Distance < keyPosArray_Int[i][1]:
                b = keyPosArray_Int[i][2]; break  
            else:
                b = 0  
       return a,b
################################################################################################################
def Reframe(df):
      # for i in df.index:             
      #       if pd.isna(df['Dia2'][i]) is True:
      #           df['Dia2'][i] = -1
      # df['Dia2'] = df['Dia1'].where(df['Dia2'] == -1, df['Dia2'])
      df['Dia2'] = df['Dia2'].fillna(df['Dia1'])
           
                
      return df    

################################################################################################################    
def Split_df(df):
     df_ext = pd.DataFrame()
     df_int = pd.DataFrame()       
     for i in df.index:
            if (df['I_E'][i] == "External"):
                 df_ext = pd.concat([df_ext, df.iloc[[i]]])
            elif (df['I_E'][i] == "Internal"):
                 df_int = pd.concat([df_int, df.iloc[[i]]])  
     df_ext.reset_index(inplace = True, drop = True)
     df_int.reset_index(inplace = True, drop = True)
     
     return df_ext, df_int
       
################################################################################################################
def Extend_dataframe(df):
      #### TOP
      df['LeftEndCoordinates'] = np.NAN
      df['LeftRadiusCoordinates'] = np.NAN
      df['LeftRadiusStartAngle'] = np.NAN
      df['LeftRadiusEndAngle'] = np.NAN
      df['LeftChamferCoordinates'] = np.NAN      
      df['RightEndCoordinates'] = np.NAN
      df['RightRadiusCoordinates'] = np.NAN
      df['RightRadiusStartAngle'] = np.NAN
      df['RightRadiusEndAngle'] = np.NAN
      df['RightChamferCoordinates'] = np.NAN
      df['GrooveCoordinates'] = np.NAN
      df['GrooveRadiusStartAngle'] = np.NAN
      df['GrooveRadiusEndAngle'] = np.NAN

      df['LeftEndCoordinates'] = df['LeftEndCoordinates'].astype('object')
      df['LeftRadiusCoordinates'] = df['LeftRadiusCoordinates'].astype('object')
      df['LeftRadiusStartAngle'] = df['LeftRadiusStartAngle'].astype('object')
      df['LeftRadiusEndAngle'] = df['LeftRadiusEndAngle'].astype('object')
      df['LeftChamferCoordinates'] = df['LeftChamferCoordinates'].astype('object')     
      df['RightEndCoordinates'] = df['RightEndCoordinates'].astype('object')
      df['RightRadiusCoordinates'] = df['RightRadiusCoordinates'].astype('object')
      df['RightRadiusStartAngle'] = df['RightRadiusStartAngle'].astype('object')
      df['RightRadiusEndAngle'] = df['RightRadiusEndAngle'].astype('object')
      df['RightChamferCoordinates'] = df['RightChamferCoordinates'].astype('object')
      df['GrooveCoordinates'] = df['GrooveCoordinates'].astype('object')
      df['GrooveRadiusStartAngle'] = df['GrooveRadiusStartAngle'].astype('object')
      df['GrooveRadiusEndAngle'] = df['GrooveRadiusEndAngle'].astype('object')


      #### BOTTOM
      df['LeftEndCoordinates1'] = np.NAN
      df['LeftRadiusCoordinates1'] = np.NAN
      df['LeftRadiusStartAngle1'] = np.NAN
      df['LeftRadiusEndAngle1'] = np.NAN
      df['LeftChamferCoordinates1'] = np.NAN      
      df['RightEndCoordinates1'] = np.NAN
      df['RightRadiusCoordinates1'] = np.NAN
      df['RightRadiusStartAngle1'] = np.NAN
      df['RightRadiusEndAngle1'] = np.NAN
      df['RightChamferCoordinates1'] = np.NAN
      df['GrooveCoordinates1'] = np.NAN
      df['GrooveRadiusStartAngle1'] = np.NAN
      df['GrooveRadiusEndAngle1'] = np.NAN

      df['LeftEndCoordinates1'] = df['LeftEndCoordinates1'].astype('object')
      df['LeftRadiusCoordinates1'] = df['LeftRadiusCoordinates1'].astype('object')
      df['LeftRadiusStartAngle1'] = df['LeftRadiusStartAngle1'].astype('object')
      df['LeftRadiusEndAngle1'] = df['LeftRadiusEndAngle1'].astype('object')
      df['LeftChamferCoordinates1'] = df['LeftChamferCoordinates1'].astype('object')     
      df['RightEndCoordinates1'] = df['RightEndCoordinates1'].astype('object')
      df['RightRadiusCoordinates1'] = df['RightRadiusCoordinates1'].astype('object')
      df['RightRadiusStartAngle1'] = df['RightRadiusStartAngle1'].astype('object')
      df['RightRadiusEndAngle1'] = df['RightRadiusEndAngle1'].astype('object')
      df['RightChamferCoordinates1'] = df['RightChamferCoordinates1'].astype('object')
      df['GrooveCoordinates1'] = df['GrooveCoordinates1'].astype('object')
      df['GrooveRadiusStartAngle1'] = df['GrooveRadiusStartAngle1'].astype('object')
      df['GrooveRadiusEndAngle1'] = df['GrooveRadiusEndAngle1'].astype('object')
      

      return df

################################################################################################################

def df_FillCoordinates(SP,df,Int_Ext):
      Left_Coordinates = []
      Right_Coordinates = []
      df_len = len(df.index)
      # k = 0
      p2 = APoint(0,0)
      # Groove_SP = APoint(0,0)
      
      for k in range(df_len):
            # Common
            TaperAngle = Taper_Angle(df['Dia1'][k],df['Dia2'][k], df['Len'][k])
            p2.x = SP.x + df['Len'][k]; p2.y = SP.y
            if Int_Ext == "External":
                  if k == 0: Prev_Dia = 0
                  else: Prev_Dia = df['Dia2'][k-1]
                  if k == df_len - 1: Next_Dia = 0
                  else: Next_Dia = df['Dia1'][k+1]
            else:
                  if k == 0: Prev_Dia = 100000
                  else: Prev_Dia = df['Dia2'][k-1]
                  if k == df_len - 1: Next_Dia = 100000
                  else: Next_Dia = df['Dia1'][k+1]

            # Left Coordinates
                  
            if df['Left_C_F'][k] == "F":        
                  df.at[k,'LeftRadiusCoordinates'], df.at[k,'LeftRadiusStartAngle'], df.at[k,'LeftRadiusEndAngle'] = Radius_Coordinates(SP,Prev_Dia, df['Dia1'][k], df['Left_Size'][k], "Left" ,TaperAngle,"Top")   
                  df.at[k,'LeftRadiusCoordinates1'], df.at[k,'LeftRadiusStartAngle1'], df.at[k,'LeftRadiusEndAngle1'] = Radius_Coordinates(SP,Prev_Dia, df['Dia1'][k], df['Left_Size'][k], "Left" ,TaperAngle,"Bottom")                     

            elif df['Left_C_F'][k] == "C":
                  df.at[k, 'LeftChamferCoordinates'] = Chamfer_Coordinates(SP,Prev_Dia, df['Dia1'][k], df['Left_Size'][k], "Left" ,TaperAngle,"Top") 
                  df.at[k, 'LeftChamferCoordinates1'] = Chamfer_Coordinates(SP,Prev_Dia, df['Dia1'][k], df['Left_Size'][k], "Left" ,TaperAngle,"Bot") 
            else:
                  df.at[k,'LeftEndCoordinates'] = End_Coordinates(SP, df['Dia1'][k],"Top")
                  df.at[k,'LeftEndCoordinates1'] = End_Coordinates(SP, df['Dia1'][k],"Bot")
                    
                     
                    

            # Right Coordinates
            if df['Right_C_F'][k] == "F":        
                  df.at[k,'RightRadiusCoordinates'], df.at[k,'RightRadiusStartAngle'], df.at[k,'RightRadiusEndAngle'] = Radius_Coordinates(p2,df['Dia2'][k], Next_Dia, df['Right_Size'][k], "Right" ,TaperAngle, "Top")   
                  df.at[k,'RightRadiusCoordinates1'], df.at[k,'RightRadiusStartAngle1'], df.at[k,'RightRadiusEndAngle1'] = Radius_Coordinates(p2,df['Dia2'][k], Next_Dia, df['Right_Size'][k], "Right" ,TaperAngle, "Bot") 
                  # df.at[k,'RightRadiusCoordinates'] = Right_Coordinates
                  # df.at[k,'RightRadiusStartAngle'] = RightStartAngle
                  # df.at[k,'RightRadiusEndAngle'] = RightEndAngle
            elif df['Right_C_F'][k] == "C":
                  df.at[k, 'RightChamferCoordinates'] = Chamfer_Coordinates(p2,df['Dia2'][k], Next_Dia, df['Right_Size'][k],"Right",TaperAngle, "Top") 
                  df.at[k, 'RightChamferCoordinates1'] = Chamfer_Coordinates(p2,df['Dia2'][k], Next_Dia, df['Right_Size'][k],"Right",TaperAngle, "Bot")
            else:
                  df.at[k,'RightEndCoordinates'] = End_Coordinates(p2, df['Dia2'][k],"Top")
                  df.at[k,'RightEndCoordinates1'] = End_Coordinates(p2, df['Dia2'][k],"Bot")


            # Groove Coordinates
            
            # Groove_SP.x = SP.x + df['Gr_Strt'][k] + df['Gr_Wid'][k] / 2; Groove_SP.y = SP.y
            df.at[k,'GrooveCoordinates'], df.at[k,'GrooveRadiusStartAngle'], df.at[k,'GrooveRadiusEndAngle'] = Groove_Coordinates(SP,df['Gr_Type'][k], df['Gr_Wid'][k], df['Dia1'][k],df['Gr_Dia'][k],df['Gr_Strt'][k],TaperAngle,df['I_E'][k],"Top")
            df.at[k,'GrooveCoordinates1'], df.at[k,'GrooveRadiusStartAngle1'], df.at[k,'GrooveRadiusEndAngle1'] = Groove_Coordinates(SP,df['Gr_Type'][k], df['Gr_Wid'][k], df['Dia1'][k],df['Gr_Dia'][k],df['Gr_Strt'][k],TaperAngle,df['I_E'][k],"Bot")
            # Increment
            SP.x = SP.x + df['Len'][k]



      return df


################################################################################################################

def Groove_Coordinates(Step_SP, Type, Width, Dia1, Groove_Dia, Groove_Start, Taper_Angle, Int_Ext,Position):    # for circular groove, ignore groove dia
      SP = APoint(0,0)
      Flag = getFlag(Position)
      SP.x = Step_SP.x + Groove_Start + Width/2; SP.y = Step_SP.y + (Dia1/2 + (Groove_Start + Width/2) * tan(Taper_Angle*pi/180))*Flag
      delta = Width/2 * tan(Taper_Angle*pi/180)
      # Step_Dia = Dia + 0
      # Depth = (Dia1 - Groove_Dia)/2
      if Int_Ext == "External": 
            Start_Angle_Top = 180 + Taper_Angle; End_Angle_Top = 0 + Taper_Angle
            Start_Angle_Bot = 0 - Taper_Angle; End_Angle_Bot = 180 - Taper_Angle
      else: 
            Start_Angle_Top = 0 + Taper_Angle; End_Angle_Top = 180 + Taper_Angle
            Start_Angle_Bot = 180 - Taper_Angle; End_Angle_Bot = 0 - Taper_Angle

      Coordinates = []
      if Type in ("O-ring","Circlip Heavy","Circlip Normal"):
                Coordinates.extend([SP.x - Width/2,SP.y - (delta)*Flag,0]) 
                Coordinates.extend([SP.x - Width/2,Step_SP.y + (Groove_Dia/2)*Flag,0]) 
                Coordinates.extend([SP.x + Width/2,Step_SP.y + (Groove_Dia/2)*Flag,0]) 
                Coordinates.extend([SP.x + Width/2,SP.y + (delta)*Flag,0])   
      else:
                
                Coordinates.extend([SP.x,SP.y,0])
                
      if Position  == "Top":
              Start_Angle = Start_Angle_Top; End_Angle = End_Angle_Top
      else:
              Start_Angle = Start_Angle_Bot; End_Angle = End_Angle_Bot

      return Coordinates, Start_Angle, End_Angle 
                
                

################################################################################################################

def DrawStep (df,Section_Type):
  SP = APoint(0,0)

  Coordinates_Hor = []
  Coordinates_Ver = []
 
  df_len = len(df.index)
  


  for i in range(df_len):
    LineType_Top, LineType_Bottom, TopFlag, BottomFlag = getLineType(Section_Type,df['I_E'][i])
    if df['I_E'][i] == "External":
            if i == 0: Prev_Dia = 0
            else: Prev_Dia = df['Dia2'][i-1]
            if i == df_len - 1: Next_Dia = 0
            else: Next_Dia = df['Dia1'][i+1]             

    else:
            if i == 0: Prev_Dia = 100000
            else: Prev_Dia = df['Dia2'][i-1]
            if i == df_len - 1: Next_Dia = 100000
            else: Next_Dia = df['Dia1'][i+1]
    
    
    # LEFT SIDE TREATMENT
    Side_Treatment(df['Left_C_F'][i],df['LeftChamferCoordinates'][i],df['LeftRadiusCoordinates'][i],df['Left_Size'][i],df['LeftRadiusStartAngle'][i],df['LeftRadiusEndAngle'][i], LineType_Top)   #Left    
    Side_Treatment(df['Left_C_F'][i],df['LeftChamferCoordinates1'][i],df['LeftRadiusCoordinates1'][i],df['Left_Size'][i],df['LeftRadiusStartAngle1'][i],df['LeftRadiusEndAngle1'][i], LineType_Bottom)
    # GROOVE TREATMENT
    Groove_Treatment(df['Gr_Type'][i],df['GrooveCoordinates'][i],df['Gr_Wid'][i],df['GrooveRadiusStartAngle'][i],df['GrooveRadiusEndAngle'][i],LineType_Top)# GROOVE
    Groove_Treatment(df['Gr_Type'][i],df['GrooveCoordinates1'][i],df['Gr_Wid'][i],df['GrooveRadiusStartAngle1'][i],df['GrooveRadiusEndAngle1'][i],LineType_Bottom)
    # RIGHT SIDE TREATMENT  
    Side_Treatment(df['Right_C_F'][i],df['RightChamferCoordinates'][i],df['RightRadiusCoordinates'][i],df['Right_Size'][i],df['RightRadiusStartAngle'][i],df['RightRadiusEndAngle'][i], LineType_Top)   #Right  
    Side_Treatment(df['Right_C_F'][i],df['RightChamferCoordinates1'][i],df['RightRadiusCoordinates1'][i],df['Right_Size'][i],df['RightRadiusStartAngle1'][i],df['RightRadiusEndAngle1'][i], LineType_Bottom)   


    ##### HORIZONTAL LINES
    Coordinates_Hor, Coordinates_Hor1 = Horizontal_Edge_Coordinates(df.iloc[i],Prev_Dia,Next_Dia)
    Horizontal_Treatment(Coordinates_Hor, LineType_Top)
    Horizontal_Treatment(Coordinates_Hor1, LineType_Bottom)


    ##### VERTICAL LINES 
    if i != len(df)-1:
      Coordinates_Ver = Vertical_Edge_Coordinates(df.iloc[i],df.iloc[i+1], Prev_Dia, Next_Dia)
      Coordinates_Ver1 = Vertical_Edge_Coordinates(df.iloc[i],df.iloc[i+1], Prev_Dia, Next_Dia)
    PL = DrawLine(Coordinates_Ver, LineType_Top) 
    PL = DrawLine(Coordinates_Ver1, LineType_Bottom) 

    # WHEN TO STOP
    if df['Dia1'][i] == "": break

################################################################################################################
def Horizontal_Edge_Coordinates(row,Prev_Dia,Next_Dia):
      Coordinates = []
      Coordinates1 = []
      PNT1 = []; PNT2 = []; PNT3 = []; PNT4 = []
      PNT1_1 = []; PNT2_1 = []; PNT3_1 = []; PNT4_1 = []     
      
      # LEFT COORDINATE
      if pd.isna(row['Left_C_F']):
            PNT1 = row['LeftEndCoordinates']
            PNT1_1 = row['LeftEndCoordinates1']
            # print (PNT1)
      elif row['Left_C_F'] == "C":
            PNT1 = row['LeftChamferCoordinates'][3:]     
            PNT1_1 = row['LeftChamferCoordinates1'][3:]   
            # print (PNT1)     
      else:
            PNT1 = (get_Radius_Coordinates(row['LeftRadiusCoordinates'], row['Left_Size'], row['LeftRadiusStartAngle'], row['LeftRadiusEndAngle'], "Left", row['Dia1'], Prev_Dia))[3:]
            PNT1_1 = (get_Radius_Coordinates(row['LeftRadiusCoordinates1'], row['Left_Size'], row['LeftRadiusStartAngle1'], row['LeftRadiusEndAngle1'], "Left", row['Dia1'], Prev_Dia))[3:]
            # print (PNT1)

      # GROOVE COORDINATES
      if pd.isna(row['Gr_Type']):    
            PNT2 = []; PNT3 = []; PNT2_1 = []; PNT3_1 = []
      elif row['Gr_Type'] == "Circular":
            PNT2 = (get_Groove_Coordinates(row['GrooveCoordinates'], row['Gr_Wid']/2, row['GrooveRadiusStartAngle'], row['GrooveRadiusEndAngle'], row['I_E']))[:3]
            PNT3 = (get_Groove_Coordinates(row['GrooveCoordinates'], row['Gr_Wid']/2, row['GrooveRadiusStartAngle'], row['GrooveRadiusEndAngle'], row['I_E']))[3:]

            PNT2_1 = (get_Groove_Coordinates(row['GrooveCoordinates1'], row['Gr_Wid']/2, row['GrooveRadiusStartAngle1'], row['GrooveRadiusEndAngle1'], row['I_E']))[:3]
            PNT3_1 = (get_Groove_Coordinates(row['GrooveCoordinates1'], row['Gr_Wid']/2, row['GrooveRadiusStartAngle1'], row['GrooveRadiusEndAngle1'], row['I_E']))[3:]
            # print (PNT1)
      else:
            
            PNT2 = row['GrooveCoordinates'][:3]
            PNT3 = row['GrooveCoordinates'][9:]

            PNT2_1 = row['GrooveCoordinates1'][:3]
            PNT3_1 = row['GrooveCoordinates1'][9:]

      # RIGHT COORDINATE
      if pd.isna(row['Right_C_F']):
            PNT4 = row['RightEndCoordinates']
            PNT4_1 = row['RightEndCoordinates1']
            # print (PNT1)
      elif row['Right_C_F'] == "C":
            # print(row['RightChamferCoordinates'])
            PNT4 = row['RightChamferCoordinates'][:3]  
            PNT4_1 = row['RightChamferCoordinates1'][:3] 

      else:
            PNT4= (get_Radius_Coordinates(row['RightRadiusCoordinates'], row['Right_Size'], row['RightRadiusStartAngle'], row['RightRadiusEndAngle'], "Right", row['Dia2'], Next_Dia))[:3]
            PNT4_1= (get_Radius_Coordinates(row['RightRadiusCoordinates1'], row['Right_Size'], row['RightRadiusStartAngle1'], row['RightRadiusEndAngle1'], "Right", row['Dia2'], Next_Dia))[:3]
            # print ("bI",PNT1)
                  
      Coordinates.extend(PNT1)
      if PNT2:
            Coordinates.extend(PNT2)
            Coordinates.extend(PNT3)
      Coordinates.extend(PNT4)

      Coordinates1.extend(PNT1_1)
      if PNT2_1:
            Coordinates1.extend(PNT2_1)
            Coordinates1.extend(PNT3_1)
      Coordinates1.extend(PNT4_1)

      return Coordinates, Coordinates1
      # if pd.isna(row['Gr_Type']):
      #       if 
                 
      # else:
      #       print ("Bye",row["Len"])

################################################################################################################            
def get_Radius_Coordinates(SP,Radius,StartAngle,EndAngle,Side,Step_Dia,Dia):
      PNT1 = []
      PNT2 = []
      Coordinates = []

      PNT1 = [SP[0]+Radius*cos(StartAngle*pi/180),SP[1]+Radius*sin(StartAngle*pi/180),0]
      PNT2 = [SP[0]+Radius*cos(EndAngle*pi/180),SP[1]+Radius*sin(EndAngle*pi/180),0]    

      
      if Side == "Left":
            if Dia>Step_Dia:
                  Coordinates.extend(PNT1)
                  Coordinates.extend(PNT2)
            else:
                  Coordinates.extend(PNT2)
                  Coordinates.extend(PNT1)
      else:
            if Step_Dia>Dia:
                  Coordinates.extend(PNT2)
                  Coordinates.extend(PNT1)
            else:
                  Coordinates.extend(PNT1)
                  Coordinates.extend(PNT2)
      return Coordinates
################################################################################################################ 
def get_Groove_Coordinates(SP,Radius,StartAngle,EndAngle,Int_Ext):
      PNT1 = []
      PNT2 = []
      Coordinates = []

      PNT1 = [SP[0]+Radius*cos(StartAngle*pi/180),SP[1]+Radius*sin(StartAngle*pi/180),0]
      PNT2 = [SP[0]+Radius*cos(EndAngle*pi/180),SP[1]+Radius*sin(EndAngle*pi/180),0]  

      if Int_Ext == "External":            
                  Coordinates.extend(PNT1)
                  Coordinates.extend(PNT2)
      else:                
                  Coordinates.extend(PNT2)
                  Coordinates.extend(PNT1)
      return Coordinates
################################################################################################################ 
def Vertical_Edge_Coordinates(row, nextrow, Prev_Dia, Next_Dia):

      Coordinates = []
      PNT1 = [] 
      PNT2 = []

      Coordinates1 = []
      PNT1_1 = [] 
      PNT2_1 = []

      # RIGHT COORDINATE
      if pd.isna(row['Right_C_F']):
            PNT1 = row['RightEndCoordinates']
            PNT1_1 = row['RightEndCoordinates1']
            
      elif row['Right_C_F'] == "C":           
            PNT1 = row['RightChamferCoordinates'][3:]  
            PNT1_1 = row['RightChamferCoordinates1'][3:]  

      else:
            PNT1= (get_Radius_Coordinates(row['RightRadiusCoordinates'], row['Right_Size'], row['RightRadiusStartAngle'], row['RightRadiusEndAngle'], "Right", row['Dia2'], Next_Dia))[3:]
            PNT1_1= (get_Radius_Coordinates(row['RightRadiusCoordinates1'], row['Right_Size'], row['RightRadiusStartAngle1'], row['RightRadiusEndAngle1'], "Right", row['Dia2'], Next_Dia))[3:]
           
      Coordinates.extend(PNT1)
      Coordinates1.extend(PNT1_1)
      # LEFT COORDINATE
      if pd.isna(nextrow['Left_C_F']):
            PNT2 = nextrow['LeftEndCoordinates']
            PNT2_1 = nextrow['LeftEndCoordinates1']
           
      elif nextrow['Left_C_F'] == "C":
            PNT2 = nextrow['LeftChamferCoordinates'][:3]  
            PNT2_1 = nextrow['LeftChamferCoordinates1'][:3]      
               
      else:
            PNT2 = (get_Radius_Coordinates(nextrow['LeftRadiusCoordinates'], nextrow['Left_Size'], nextrow['LeftRadiusStartAngle'], nextrow['LeftRadiusEndAngle'], "Left", nextrow['Dia1'], Prev_Dia))[:3]
            PNT2_1 = (get_Radius_Coordinates(nextrow['LeftRadiusCoordinates1'], nextrow['Left_Size'], nextrow['LeftRadiusStartAngle1'], nextrow['LeftRadiusEndAngle1'], "Left", nextrow['Dia1'], Prev_Dia))[:3]

      Coordinates.extend(PNT2)
      Coordinates1.extend(PNT2_1)
      return Coordinates, Coordinates1
           
      
def Prev_Next_Dia(Int_Ext,Dia1, Dia2, index, i):   #####NOT IN USE###########
      if Int_Ext == "External":
            if i == 0: Prev_Dia = 0
            else: Prev_Dia = Dia2
            if i == index - 1: Next_Dia = 0
            else: Next_Dia = Dia1
      else:
            if i == 0: Prev_Dia = 100000
            else: Prev_Dia = Dia2
            if i == index - 1: Next_Dia = 100000
            else: Next_Dia = Dia1
      return Prev_Dia, Next_Dia
################################################################################################################
def Side_Treatment(C_F,ChamferCoordinates,RadiusCoordinates,Size,RadiusStartAngle,RadiusEndAngle,LineType):
      if C_F == "C":
             PL = DrawLine(ChamferCoordinates, LineType)
      elif C_F == "F":
             PL = DrawArc(RadiusCoordinates,Size,RadiusStartAngle,RadiusEndAngle,LineType) 

################################################################################################################
def Groove_Treatment(Gr_Type,GrooveCoordinates,Gr_Wid,GrooveRadiusStartAngle,GrooveRadiusEndAngle,LineType):
        
        if Gr_Type == "Circular":
            PL = DrawArc(GrooveCoordinates, Gr_Wid/2, GrooveRadiusStartAngle, GrooveRadiusEndAngle,LineType)
        elif Gr_Type in ("O-ring","Circlip Heavy","Circlip Normal"):
            PL = DrawLine(GrooveCoordinates, LineType)

################################################################################################################
def Horizontal_Treatment(Coordinates_Hor, LineType):
    if len(Coordinates_Hor) == 6:
      PL = DrawLine(Coordinates_Hor, LineType)
    else:
      PL = DrawLine(Coordinates_Hor[:6], LineType)  
      PL = DrawLine(Coordinates_Hor[6:], LineType)  

################################################################################################################
def getLineType(Type):
       if Type == "Fullsection":
              x= 0
       elif Type == "Nosection":
              x = 1
       else:
              x = 3

       return LineType_Top, LineType_Bottom, Ex_LineType_Top, Ex_LineType_Bottom       


################################################################################################################