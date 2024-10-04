sentimientos = open("Sentimientos.txt") 
valores = {} 
for linea in sentimientos: 
    termino, valor = linea.split("\t") 
    valores[termino] = int(valor) 
print (valores.items() )