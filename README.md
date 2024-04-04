## Excel para SQL

Função em javascript que converte a sintaxe das funções E, OU e NÃO do Excel para a sintaxe SQL.
<br />
Uso:
<br />
let texto = ExcelParaSQL({
<br />
  &nbsp;&nbsp;&nbsp;&nbsp;conteudo : texto, // string obrigatória
<br />
  &nbsp;&nbsp;&nbsp;&nbsp;separador : ",", // char opcional com separador. Padrão ","
<br />
  &nbsp;&nbsp;&nbsp;&nbsp;campos : [] // array de strings com os campos que serão permitidos. Padrão array vazio, permite todos
<br />
});
