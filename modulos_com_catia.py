
import os 
from win32com.client import Dispatch

def conectar_com():
    CATIA = Dispatch('CATIA.Application')
    CATIA.Visible = True
    return CATIA

def crear_documento(nombre,objeto_com):
    parte = objeto_com.Documents.Add('Part')
    product1 = parte.GetItem("Part1")
    product1.PartNumber = nombre
    return parte

def crear_circunferencia(coord_ax1,coord_ax2,radio,sketch):
    newcircle = sketch.Factory2D.CreateClosedCircle(coord_ax1, coord_ax2, radio)
    return newcircle

def abrir_documento(ruta,objeto_cad):
    partDocument1 = objeto_cad.Documents.Open(ruta)
    return partDocument1



if __name__ == "__main__":
    objeto = conectar_com()
    parte = abrir_documento("C:\\Users\\sandu\\Desktop\\modulos_com_catia\\anillo.CATPart",objeto)
    #se navega a traves de la infraestructura del com para acceder a los atributos del modelo
    parte1 = parte.Part
    geometrias = parte1.GeometricElements
    figura = geometrias.Item("Circle.1")
    print(figura.Radius)
    figura2 = geometrias.Item("Circle.2")
    print(figura2.Radius)

    



    """
    referencia_parte = crear_documento('anillo',objeto)
    Xcoord=100
    Ycoord=100
    Zcoord=100
    NewPoint = objeto.ActiveDocument.Part.HybridShapeFactory.AddNewPointCoord(Xcoord, Ycoord, Zcoord)
    Mainbody = objeto.ActiveDocument.Part.MainBody
    Mainbody.InsertHybridShape(NewPoint)
    AxisXY = objeto.ActiveDocument.Part.OriginElements.PlaneXY
    Referenceplane = objeto.ActiveDocument.Part.CreateReferenceFromObject(AxisXY)
    Referencepoint = objeto.ActiveDocument.Part.CreateReferenceFromObject(NewPoint)
    NewPlane = objeto.ActiveDocument.Part.HybridShapeFactory.AddNewPlaneOffsetPt(Referenceplane, Referencepoint)
    Mainbody.InsertHybridShape(NewPlane)    
    sketches1 = objeto.ActiveDocument.Part.Bodies.Item("PartBody").Sketches
    reference1 = referencia_parte.part.OriginElements.PlaneXY
    NewSketch = sketches1.Add(reference1)
    objeto.ActiveDocument.Part.InWorkObject = NewSketch
    NewSketch.OpenEdition()
    newcircle = crear_circunferencia(0, 0, 20.215757,NewSketch)
    newellipse = crear_circunferencia(0, 0, 10.069477,NewSketch)
    NewSketch.CloseEdition()    
    #grosor del bloque
    LengthBlock=11
    NewBlock = objeto.ActiveDocument.Part.ShapeFactory.AddNewPad (NewSketch, LengthBlock)
    objeto.ActiveDocument.Part.Update()
    product1 = referencia_parte.GetItem("Rotor_dinamico")
    """
    