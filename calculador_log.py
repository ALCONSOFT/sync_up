import logging
logging.basicConfig(level=logging.DEBUG)

class calculadora():

  def __init__():
    print("\nCalculadora encendida.")

  def sumar( a=0, b=0):
    print("Suma a={} b={} a+b={}".format(a,b,a+b))

  def restar( a=0, b=0):
    print("Resta a={} b={} a-b={}".format(a,b,a-b))

  def multiplicar( a=0, b=0):
    print("Multiplicación a={} b={} a*b={}".format(a,b,a*b))

  def dividir( a=0, b=1):
    print("División a={} b={} a/b={}".format(a,b,a/b))

calc = calculadora
calc.sumar(12,4)
calc.restar(12,4)
calc.multiplicar(12,4)
calc.dividir(12,4)