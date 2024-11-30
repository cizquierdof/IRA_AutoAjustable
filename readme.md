# IRA  autoajustable a la muestra
Libro xcel habilitado para macros. Contiene módulos VBA que generan las hojas Perceptible, Operable, Comprensible y Robusto, además de la de Resultados.
Las hojas de los principios generan una tabla por cada criterio que tiene un nómero de filas igual al de la muestra, por lo que ya no está limitado a 35.

La hoja de resultaodos se genera con fórmulas que trasladan los resultados a dos zonas (no son tablas) separadas para los criterios A y AA con filas de resultados.

se puede utilizar VSCode para editar los módulos y que se reflejen los cambios en el editor VBA de Excel peero no funciona muy bien. Para que esto funciones hacer lo siguiente:

## Arrancar el entorno de VBA
se necesita:
- Python
- xlwings
- watchgod
- Libro excel habilitado para macros y permitir el entorno VBA en las opciones del centro de confianza de Office.



Una vez instaldo el entorno para arrancar se abre la carpeta del proyecto en VSCode, se añade el módulo que queremos editar en el libro Excel y se guarda. Después se ejecuta el comando:

```
xlwings vba edit
```
si solo hay un Excel o:
```
xlwings vba edit name.xlsm
```





