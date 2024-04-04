## Excel para SQL

Função em javascript que converte a sintaxe das funções E, OU e NÃO do Excel para a sintaxe SQL.
<br />
Uso:
<br />
let texto = ExcelParaSQL({
<br />
  conteudo : texto, // string obrigatória
<br />
  separador : ",", // char opcional com separador. Padrão ","
<br />
  campos : [] // array de strings com os campos que serão permitidos. Padrão array vazio, permite todos
<br />
});
