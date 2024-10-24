## Excel para SQL

Função em javascript que converte a sintaxe das funções E, OU e NÃO do Excel para a sintaxe SQL.
<br />
Uso:
<br />
ExcelParaSQL({
<br />
  &nbsp;&nbsp;&nbsp;&nbsp;conteudo : texto, // string obrigatória
<br />
  &nbsp;&nbsp;&nbsp;&nbsp;separador : ",", // char opcional com separador. Padrão ","
<br />
  &nbsp;&nbsp;&nbsp;&nbsp;campos : [] // array de strings com os campos que serão permitidos. Padrão array vazio, permite todos
<br />
  &nbsp;&nbsp;&nbsp;&nbsp;tipos : [] // array de booleanos que são verdadeiros caso o campo correspondente seja uma string
<br />
});
<br /><br />
A função retorna a chave "sql", que é a condição SQL respectiva e "campos" que são as colunas que a função encontrou na string.
