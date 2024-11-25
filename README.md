# node-red-contrib-xlsx-to-json-2
This node is a fork from [node-red-contrib-xlsx-to-json](https://www.npmjs.com/package/node-red-contrib-xlsx-to-json). 
It resolves a bug to accept diff filepaths in multiple executions.


### Original Read me
------------------------ 
Castellano

Nodo para convertir en JSON un archivo xlsx, pudiendo seleccionar una pestaña o un rango específico del excel.

Los parámetros que necesita para funcionar son:

msg.filepath: La ruta donde se encuentra el archivo xlsx

msg.rangecell: Opcional. El rango de celdas que se leeran del excel. Ejemplo -> "A2:M6"

msg.columkey: Opcional. La cabecera para cada columna que se va a leer. Es un objeto. Ejemplo -> {"A":"Titulo","B":"Descripción"}

msg.sheet: El nombre de la hoja del libro que se van a leer. Ejemplo -> Hoja 1

--------------------------

English

Node to convert an xlsx file into JSON, being able to select a tab or a specific range of excel.

The parameters it needs to work are:

msg.filepath:The path where the xlsx file is located

msg.rangecell: Optional. The range of cells to be read from excel. Example -> "A2:M6"

msg.columkey: Optional. The header for each column to be read. Its an object. Example -> {"A":"Titulo","B":"Descripción"}

msg.sheet: The name of the book sheets to be read. Example -> Sheet 1
