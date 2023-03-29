"""APUNTES

CLASE Y METODOS
SUPERCLASE
- Class Nom_clase:
    atributoX = x -----------> es un atributo de clase, queda definido para todas las instancias de la misma forma

    @classmethod
    def metodo_de_clase(cls):  -----------> metodo de clase. solo puede ser accedido por la clase y no por instancias de la misma. se usa para modificar
        return cls.Nom_clase.atributoX      o referenciar la clase en si. implementacion Nom_clase.metodo_de_clase

    def __init__ (self, a1, a2, etc):
        self.atributo1 = a1 -----------> es un atributo especifico de la instancia/objeto (queda indicado por el self)
        self.atributo2 = a2
        self.atributoetc = etc

    def get_atributo1 (self):
        return self.atributo1

    def set_atributo1 (self, nuevoa1):
        self.atributo1 = nuevoa1

    @staticmethod
    def sumar5(x):   ---------> funcion pura, no se necesita crear una instancia para usar el metodo. no puede acceder a atributos y metodos de clase
        return x+5              se crea dentro de clases para armar paquetes de funciones y facilitar su importacion

SUBCLASE (hereda caracteristicas (atributos y metodos) de superclase al pasarle como argumento el nombre de la clase)
- Class Sub_nom_clase(nom_clase):
    def __init__(self, a1, a2, "nuevo_a3", etc):
        super().__init__(a1, a2, etc)
        self.atributo3 = nuevo_a3

    def metodo_sub(self):
        blablabla

#MODIFICAR ATRIBUTO DE CLASE
clase.atributo = nuevo_valor
ej: nom_clase.atributoX = z

#F-STRINGS
string = f"hola mi nombre es {variable}"

#SLICE OPERATOR (a y b son posiciones indice. funciona con strings)
nuevalista= lista[a(empieza):b(termina pero no incluido):c(saltos)] -------> genera una nueva lista

nuevalista= lista[a:] --------> empieza en a hasta el final

nuevalista= lista[:b] -------> empieza en el principio y termina en b

nuevalista= lista[::-1] ------> da vuelta la lista

#TUPLAS va entre ()
son listas inmutables. se las puede recorrer y acceder pero no modificar

#SETS va entre {}
conjuntos no ordenados. no tiene index position, lo unico que interesa es si un elemento existe o no. contiene elementos unicos no repetidos
Para chequear si un elemento pertenece al set
x in set --------> arroja boolean
set()-----> crea un set vacio

#DICCIONARIOS va entre {}
conjunto de variables con valores asignados
ej: dicc={"val1": 4,"val2": [1,2,3]}

para obtener valor de atributo de diccioanrio
dicc["val"]

para obtener valor de un atributo de diccionario pero arroja un msj si no lo encuentra
dicc.get("valZ","no se encontro valZ")

para agergar campo al diccionario
dicc["nuevo_campo"]= valor

para modificar varios campos al mismo tiempo
dic.update("val1": 10, "val2":["perro","gato"])

para recorrer diccionarios se usa el metodo item()
ej: for key, value in mi_dic.items():
        print (key, value)

#LAMBDA FUNCTIONS
funciones anonimas. no se definen
lambda nom_arbitrario : nom_arbitrario + 3

para utilizarlas se les puede asingar un nombre y luego utilizarlas como una funcion regular
g = lambda x: x+1
g(1)
=2

OTRAS UTILIDADES
ej: x=[]
nueva x= map(lambda i: i + 2, x) ------> suma a cada elemento de la lista         LOS METODOS MAP, FILTER Y REDUCE NO DEVUELVEN LISTAS. HAY QUE PARSEARLOS
nueva x= filter(lambda i: i % 2 == 0,x) -------> deja solo los multiplos de 2
nueva x= reduce(lambda x, y: a + b, x) ------> suma el 1 y 2 elemento de la lista, al resultado le suma el 3 y asi sucesivamente

#METODOS DE LOS STRING
.upper() = hace mayusc
.lower() = hace min
.capitalize() = hace la primer letra mayusc
.count("x") = se fija si x esta en el string y arroja cantidad de veces que se repite. case sensitive

#OPERACIONES SOBRE SETS
-union
-intersection
-difference -----> dados los conjuntos a y b. la interseccion de a con b es un subconjunto que contiene los elementos de a que "no" estan en b.
                   es decir, los a "puros".

#FUNCION ISINSTANCE
sirve para comprobar si un objeto es de determinada clase, arroja un boolean.
ej: class xyz:
        pass

    a =xyz()
    isinstance(a, xyz)-----> TRUE

#METODOS DE LISTAS
.insert(index_pos, values_to_add)------> si el valor es un lista inserta la lista entera en la posicion indicada / similar a .append
.extend(values_to_add)------> agerga los elementos del conjunto como nuevos elemntos individuales de la lista
.pop------> elimina el ultimo elemento de la lista y imprime por lo q se puede almacenar en una variable
.reverse-----> invierte la lista
.sort(reverse=True/False)-----> ordena alfabaticamente lista o de menor a mayor
estos metodos alteran la lista original
.sorted()---->igual que sort pero crea una nueva version de la lista sin modificar la original
.index(value)-----> nos da el indice del argumento pasado

#ENUMERATE EN CICLOS FOR
sirve para correr a traves de una lista y traer tanto el numero indice como el elemento
ej: for index, i in enumerate(lista, start=x)

#JOIN
sirve para escribir los elementos de una lista como un string separados por algun caracter.
ej: lista=["pablo","pedro","juan"]
"char_separador".join(lista)

#SPLIT
sirve para separar un string en una lista dado un caracter que sirve como limitador
ej: string= "hola-como-estas"
string.split("-")

#ZIP
para recorrer listas en simultaneo
list1 = ["a", "b", "c"]
list2 = [1, 2, 3, 4]
zip_object = zip(list1, list2)
for element1, element2 in zip_object:
print(element1, element2)

#COMPREHENSIONS
nueva_lista= [x+3 for x in lista]

#MANEJO DE TXTS
se utiliza un content manager "with" para trabajar con txt. los text son interpretados como listas de lineas

with open("dir.txt", "r") as file: ---- lee
with open("dir.txt", "w") as file: ---- escribe
with open("dir.txt", "a") as file: ---- append, escribe al final

METODO DEL LECTOR
file.readline(n) lee la linea n
file.tell() dice en que linea del texto esta parado el lector
file.seek(n) envia al lector a la linea n

METODO PARA ESCRIBIR
file.write("texto") no salta linea
print("texto", file=file) salta linea

PARA LEER Y ESCRIBIR COMO DICS - hay que importar modulo csv
LEER
with open("dir.txt", "r") as file:
dic_reader = csv.DictReader(file)

ESCRIBIR
with open("dir.txt", "w") as file:
dic_writer= csv.DictWriter(file, fieldnames= list_fieldnames, delimiter=";") - hay que pasarle un argumento con los headers
                                                                                como lista list_fieldnames = ["nombre", "apellido"].
                                                                                el delimiter establece por cual carecter se separaran
                                                                                las columnas

dic_writer.writeheader() escribe los header

dic_writer.writerow() escribe una linea

#CSM

#GENERADORES
YIELD
Palabra clave que se utiliza para pausar una funcion.
Funciona como en return, en el sentido que luego de declararlo puedo detallar un valor que deseo que devuelva. Por ej:

deg gen():
    for i in range(20):
        yield i

g=gen()

next(g) hace que la funcion se ejecute hasta el prox yield o devuelva el siguiente valor

#FORMAT STRING AND PLACEHOLDERS
"la variable vale {variable_A.x} y es menor que {variable_B.+}".format(variable_A=10, variable_B=20) - de acuerdo a la notacion al final de
la variable se le puede dar distintos formatos al texto

%S - una forma alternativa de meter variables dentro de un string es con %s (la s se puede remplazar por otras letras para distintos efectos)

Ej: "string %s, string %s, string %s" % (1,2,3)

#PANDAS
declaraciones de dataframe
df = df.read_excel("dir",sheet_name="sheet" ,index_col="col") ---- index col indica si quiero utilizar alguna columna como indice para las filas

datos estadsiticos
df.describre - arroja datos estadisticos descriptivos

manejo de campos vacios
df.dropna() - elimina filas con datos vacios na
df.fillna(n) - llenas las datos vacions/na con n
df.fillna({"col1":0, "col2":"not found"}) - llena datos vacios pero especificando el fill por columna

definir el indice de un dataframe a partir de una columna
df.set_index("col", inplace="True")

convertir dataframe a lista
lista=df.values

metodos de filtrados
df.head(n) - muestras las primeras n filas
df.tail(n) - muesta las ultimas n filas
-seleccion columnas
df["col1"] - invoca la col1
df[["col1","col2"]] - invoca col1 y col2
df["col1":"coln"] - invoca desde la col1 hasta la coln
-seleccion filas
df.iloc[0:n] - seleccion filas por numero integer. puedo seleccionar rangos
df.loc[] - selecciona por indice asignado
-seleccion simultanea filas y cols
df.iloc[1:5,0:3] - seleccion multiple por integers. ej desde las filas 1 a 5, con columnas de 0 a 3
df.loc[[],[]]  - seleccion multiple por indices asignados (nombres). permite pasarle una lista de booleans, ej:         PRINCIPAL METODO DE SELECCION
                 df.loc[df["col1"]="condicion"] -> df["col1"]="condicion" es una lista de posicion indice/boolean.
                 loc busca los indices trues y trae esas filas enteras (si no se especifica cuales)
                 ej2:df.loc[df["col1"]="condicion",["col2","col3"]]
                 ej3:df.loc[0:2,"hola":"puto"]

-filtrados condicionales
condicion simple
df[df["col"] == x ]
multiples condiciones
df[(df["col1"]<x)&(df["col2"]>y)]
condiciones con string
df[df["col1"]="string"] donde se cumpla estrictamente ese string
df[df["col1"].str.contains("string") donde contenga en alguna parte "string"

agregar columna
df["nueva_col"]= valor asignado

agregar columna personalizada con funcion sobre un sola columna
df["nueva_col"]= df["col_que_sufrira_el_efecto].apply(funcion) - la funcion va sin (). el argumento de la funcion va a ser ocupado por la fila
                                                                 ej: def funcion(arg):
                                                                        nuevo_valor = arg+2
                                                                        return nuevo_valor

agregar columna personalizada con funcion nutriendose de varias cols
df["nueva_col"]= df.apply(funcion, axis=1) - la funcion va sin (). al ser varias cols involucradas, el argumento en la funcion debera ser acompañados por
                                     argumento["col"]. ej: def funcion(arg):
                                                               nuevo_valor=arg["col1"]+arg["col2"]
                                                               return nuevo_valor

tambien se pueden aplicar funciones lambda sobre las columnas
df["col1"]=df["col1"].apply(lambda x: x+100)

agrupar filas
df.groupby("col1").mean() - agrupa todas las filas por los valores en col1 (deberia ser una variable categorica) y en las demas cols calcula una media

df.groupby("col1").agg({
    "col2":"sum",
    "col3":"max"
}) - agrupa por los valores de col1 y en col2 calcula suma y en col3 el max

"""






