/*
EXCEL PARA SQL © 2024
Desenvolvido por Reynolds Costa, no Notepad++
O uso é permitido; a comercialização, proibida.
*/

function ExcelParaSQL(obj){let _campos=[],_tipos=[],contagem=function(e,o,r,n,t){void 0===t&&(t=0);let a="",i=0,f=0;for(let c=0;c<e.length;c++){let s=e.substring(c,c+1);o.indexOf(s)>-1&&i++,r.indexOf(s)>-1&&f++}return i!=f+t&&(a=n),a},like=function(e,o){let r=function(e,o){let r=!0;o=o.trim().replace("(","").replace(")","");try{-1==e.indexOf(o)&&(console.warn('O campo "'+o+'" n\xe3o foi encontrado e foi removido'),r=!1)}catch(n){}return r},n=function(e){let o=/^(\d{2})\/(\d{2})\/(\d{4})$/,r=e.match(o);if(!r)return!1;let n=parseInt(r[1],10),t=parseInt(r[2],10)-1,a=parseInt(r[3],10),i=new Date(a,t,n);if(i.getFullYear()!==a||i.getMonth()!==t||i.getDate()!==n)return!1;let f=[a,t+1,n];for(let c=0;c<3;c++)f[c]=f[c].toString().padStart(2,"0");return"'"+f.join("-")+"'"},t=[];return e.forEach(e=>{if("object"==typeof e&&(e=e[0]),e.indexOf("<")>-1||e.indexOf(">")>-1){let a=e.indexOf("<")>-1?"<":">",i=e.split(a);a=" "+a+" ";let f=n(i[1]),c=i[1].replace("/","").replace("(","").replace(")",""),s=!1;c==parseFloat(c)||f?f&&(i[1]=f):(s=!0,console.error("Os operadores >, >=, < e <= s\xf3 est\xe3o dispon\xedveis para datas e valores num\xe9ricos")),!s&&r(o,i[0])&&t.push((i[0]+a+i[1]).replace("> /",">= ").replace("< /","<= "))}else if(e.indexOf("=")>-1){let d=e.split("=");r(o,d[0])&&t.push(parseInt(d[1])==d[1]?d[0]+"="+d[1]:d[0]+" LIKE '%"+d[1].split(" ").join("%")+"%'")}else t.push(e)}),t},eColuna=function(e){if(!e||"string"!=typeof e||(e=e.trim(),!/^[a-zA-Z_$]/.test(e)))return!1;for(let o=1;o<e.length;o++)if(!/^[a-zA-Z0-9_$]/.test(e[o]))return!1;return!0},todos=function(e,o,r){if(e){for(e=e.toString();e.indexOf("'")>-1;)e=e.replace("'","");let n=[];for(let t=0;t<o.length;t++)(!r[t]&&parseInt(e)==e||r[t])&&n.push(r[t]?o[t]+" LIKE '%"+e+"%'":o[t]+"="+e);return[n.join(" OR "),o]}return["1",o]};if(this.e=function(e,o){return void 0===o&&(o=e,e=null),"("+like(o,e).join(" AND ")+")"},this.ou=function(e,o){return void 0===o&&(o=e,e=null),"("+like(o,e).join(" OR ")+")"},this.nao=function(e,o){return void 0===o&&(o=e,e=null),"NOT("+like([o],e)[0]+")"},void 0!==obj){let separador=",",erro="",retorno=!1,_campos_ret=[];if("object"==typeof obj){let aviso="",aux=!1;for(x in obj)-1==["conteudo","separador","campos","tipos"].indexOf(x)&&(aviso='A chave "'+x+'" n\xe3o \xe9 conhecida');aviso&&console.warn(aviso),"string"!=typeof obj.conteudo&&(erro='A chave "conte\xfado" deve ser uma string'),-1==["undefined","string"].indexOf(typeof obj.separador)?aux=!0:"string"==typeof obj.separador&&1!=obj.separador.length&&(aux=!0),erro||(aux?erro='A chave "separador", se declarada, deve ser uma string de tamanho 1':void 0===obj.separador||(separador=obj.separador)),erro||(-1==["undefined","object"].indexOf(typeof obj.campos)?aux=!0:"object"==typeof obj.campos&&obj.campos.forEach(e=>{"string"==typeof e&&eColuna(e)||(aux=!0)}),aux?erro='A chave "campos", se declarada, deve ser um array de colunas SQL':void 0===obj.campos||(_campos=obj.campos)),erro||(-1==["undefined","object"].indexOf(typeof obj.tipos)?aux=!0:"object"==typeof obj.tipos&&obj.tipos.forEach(e=>{}),aux?erro='A chave "tipos", se declarada, deve ser um array de booleanos, verdadeiros quando o campo correspondente for uma string':void 0===obj.tipos||(_tipos=obj.tipos)),_tipos.length!=_campos.length&&(erro='Os arrays "tipos" e "campos" devem ter o mesmo tamanho, pois s\xe3o correspondentes')}else erro="Par\xe2metro n\xe3o declarado corretamente";if(!erro){var texto=obj.conteudo.toLowerCase();if(","!=separador)for(;texto.indexOf(separador)>-1;)texto=texto.replace(separador,",");let funcoes=texto.indexOf("e(")>-1||texto.indexOf("ou(")>-1||texto.indexOf("nao(")>-1;var pseudocodigo=funcoes||texto.indexOf(",")>-1||texto.indexOf("<")>-1||texto.indexOf("=")>-1||texto.indexOf(">")>-1;for(!funcoes&&pseudocodigo&&(texto="e("+texto+")");texto.indexOf(">=")>-1;)texto=texto.replace(">=",">/");for(;texto.indexOf("<=")>-1;)texto=texto.replace("<=","</");!funcoes||(erro=contagem(texto,"(",")","N\xfamero desigual de par\xeanteses"))||(erro=contagem(texto,"<=>",",","Termos n\xe3o declarados corretamente",1))}if(!erro){texto=texto.replace(/\(/g,"(['").replace(/\)/g,"'])").replace(/,/g,"','");let substituir={",'(['":",","'ou([":"ou(['","'e([":"e(['","'nao([":"nao(['","])'":"])","'ou":"ou","'e":"e","'nao":"nao","''":"'"," ":"","([ou":"(['ou","([e":"(['e","([nao":"(['nao"};for(x in substituir)for(;texto.indexOf(x)>-1;)texto=texto.replace(x,substituir[x]);texto=texto.replace(/e\(/g,"new ExcelParaSQL().e(").replace(/ou\(/g,"new ExcelParaSQL().ou(").replace(/nao\(/g,"new ExcelParaSQL().nao(");try{let texto_formatado=eval(texto),poss_col=[],poss_col2=[];if(["<","=",">"].forEach(e=>{texto.split(e).forEach(e=>{(e=e.split("[")).forEach(e=>{poss_col.push(e)})})}),poss_col.forEach(e=>{if(e.indexOf("'")>-1){let o=e.split("'");o=o[o.length-1],eColuna(o)&&-1==poss_col2.indexOf(o)&&(_campos.indexOf(o)>-1||!_campos.length)&&poss_col2.push(o)}}),texto=eval(texto.replace(/e\(/g,"e(['"+poss_col2.join("','")+"'],").replace(/ou\(/g,"ou(['"+poss_col2.join("','")+"'],").replace(/nao\(/g,"nao(['"+poss_col2.join("','")+"'],")),poss_col2.forEach(e=>{texto.indexOf(e)>-1&&_campos_ret.push(e)}),!pseudocodigo&&_campos.length){let aux=todos(texto,_campos,_tipos);texto=aux[0],_campos_ret=aux[1]}}catch(err){if(!pseudocodigo&&_campos.length){let aux=todos(texto,_campos,_tipos);texto=aux[0],_campos_ret=aux[1]}else erro="Erro desconhecido"}}if(erro)console.error(erro);else{for(texto=texto.toString();texto.indexOf("()")>-1;)texto=texto.replace("()","1");for(;texto.indexOf("  ")>-1;)texto=texto.replace("  "," ");retorno={sql:encodeURIComponent(texto),campos:_campos_ret}}return retorno}}