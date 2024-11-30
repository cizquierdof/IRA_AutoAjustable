# Arrancar el entorno de VBA
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
para uno específico




