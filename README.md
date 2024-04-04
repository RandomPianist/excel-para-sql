## Excel para SQL

Função em javascript que converte a sintaxe das funções E, OU e NÃO do Excel para a sintaxe SQL.

Uso:

let texto = ExcelParaSQL({
  conteudo : texto, // string obrigatória
  separador : ",", // char opcional com separador. Padrão ","
  campos : [] // array de strings com os campos que serão permitidos. Padrão array vazio, permite todos
});
