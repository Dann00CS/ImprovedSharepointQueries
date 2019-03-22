# ImprovedSharepointQueries
List of actions to make a sharepoint query more efficient

- SelectStaffOption:
  
Si estamos en un formulario de edicion, y dentro hay un combo que tira de Sharepoint, al entrar, se debe hacer una llamada, que pueda recuperar el ComboItemDTO que corresponde con el ID que permite elegir una option u otra.
  
 - SearchFirstXStaffRows:
 
Si estamos ante cualquier formulario, ya sea de creacion o de edicion, y tenemos un combo que tira de Sharepoint, el cual tiene un search input, cada vez que introduzcamos ahi algo, se debe buscar en Sharepoint, en busca de un elemento que contenga las letras escritas
