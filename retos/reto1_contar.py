"""Reto 1 · Contar los elementos de la tabla de forma eficiente.

    py -3 retos/reto1_contar.py
Cuenta, SIN depurar todavía, cuántas veces aparece cada id, código y categoría de la hoja
«Materiales». Compara dos algoritmos de recuento: ingenuo (listas) y con diccionario.
"""
import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "herramienta"))
import configuracion
from lector_excel import leer_hoja
from utiles import barras_texto, escribir_csv, guardar_texto, tabla_texto


def contar_ingenuo(valores):
    """Recuento con listas: para cada valor distinto, recorre TODA la lista.
    Devuelve (recuento, comparaciones). Coste: unas n × k comparaciones (k = valores distintos)."""
    comparaciones = 0
    distintos = []
    for v in valores:                       # 1. lista de valores distintos
        esta = False
        for d in distintos:
            comparaciones += 1
            if d == v:
                esta = True
                break
        if not esta:
            distintos.append(v)
    recuento = {}
    for d in distintos:                     # 2. para cada distinto, contar en toda la lista
        n = 0
        for v in valores:
            comparaciones += 1
            if v == d:
                n += 1
        recuento[d] = n
    return recuento, comparaciones


def contar_diccionario(valores):
    """Recuento con un diccionario: una sola pasada. Devuelve (recuento, pasos). Coste: n."""
    recuento = {}
    pasos = 0
    for v in valores:
        recuento[v] = recuento.get(v, 0) + 1
        pasos += 1
    return recuento, pasos


def repetidos(recuento):
    """Valores que aparecen más de una vez (sin contar el vacío)."""
    return sorted(v for v, cantidad in recuento.items() if v != "" and cantidad > 1)


def pruebas():
    assert contar_diccionario(["a", "b", "a"])[0] == {"a": 2, "b": 1}          # presente
    assert contar_diccionario(["a", "b", "a"])[0].get("z", 0) == 0               # ausente
    assert contar_diccionario([]) == ({}, 0)                                     # vacío
    assert contar_ingenuo(["a", "b", "a"])[0] == contar_diccionario(["a", "b", "a"])[0]
    assert repetidos({"a": 2, "b": 1, "": 5}) == ["a"]


if __name__ == "__main__":
    pruebas()
    configuracion.comprobar()
    materiales = leer_hoja(configuracion.RUTA_EXCEL, "Materiales")
    n = len(materiales)

    filas = []
    for columna in ("id_material", "codigo_material", "Descripcion_material_hqsmlg"):
        valores = [m[columna] for m in materiales]
        r_ing, comp = contar_ingenuo(valores)
        r_dic, pasos = contar_diccionario(valores)
        assert r_ing == r_dic                                 # los dos algoritmos coinciden
        filas.append([columna, len([v for v in r_dic if v != ""]), r_dic.get("", 0),
                      len(repetidos(r_dic)), comp, pasos])

    # Medición con la columna de códigos repetida 50 veces (para que el tiempo se note)
    grande = [m["codigo_material"] for m in materiales] * 50
    inicio = time.perf_counter()
    contar_ingenuo(grande)
    t_ing = time.perf_counter() - inicio
    inicio = time.perf_counter()
    contar_diccionario(grande)
    t_dic = time.perf_counter() - inicio

    categorias = contar_diccionario([m["Descripcion_material_hqsmlg"] for m in materiales])[0]
    top = sorted(categorias.items(), key=lambda par: -par[1])[:10]      # de más a menos
    lineas = ["# Reto 1 · Recuento de la hoja Materiales (sin depurar)", "",
              "Registros: %d" % n, ""]
    lineas += tabla_texto(["Columna", "Valores distintos", "Vacíos", "Repetidos",
                           "Comparaciones (ingenuo)", "Pasos (diccionario)"], filas)
    lineas += ["", "Tiempo con %d valores: ingenuo %.4f s · diccionario %.4f s" % (len(grande), t_ing, t_dic), "",
               "Códigos repetidos: %s" % (repetidos(contar_diccionario([m["codigo_material"] for m in materiales])[0]) or "ninguno"),
               "Ids repetidos: %s" % (repetidos(contar_diccionario([m["id_material"] for m in materiales])[0]) or "ninguno"),
               "", "```"]
    lineas += barras_texto("Las 10 categorías más frecuentes (tal como están escritas)",
                           [("«%s»" % c if c != "" else "(vacía)", v) for c, v in top])
    lineas += ["```"]
    print("\n".join(lineas))
    guardar_texto(os.path.join(configuracion.RESULTADOS, "reto1_recuento.md"), lineas)
    escribir_csv(os.path.join(configuracion.RESULTADOS, "reto1_categorias.csv"), ["categoria", "materiales"],
                 sorted(categorias.items(), key=lambda par: -par[1]))
