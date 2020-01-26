class Factura:
    empresa = ""
    numerofact = ""
    descripcion = ""
    cantidad = 0
    cancelada = False
    gasto = False
    cell=""
    def __init__(self):
        empresa=""
        numerofact=""
        descripcion=""
        cantidad=0
        cancelada=False
        gasto=False
    def __init__(self,empre,numeroF,desc,cantid,cancel,pgasto,cell):
        empresa=empre
        numerofact=numeroF
        descripcion=desc
        cantidad=cantid
        cancelada=cancel
        gasto=pgasto
        cell=cell
class ManagerFacturas:
    def __init__(self):
        CobroPago=[]
        CobroPendiente=[]
        FacturaPaga=[]
        FacturaPendiente=[]





