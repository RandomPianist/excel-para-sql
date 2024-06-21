/*
EXCEL PARA SQL © 2024
Desenvolvido por Reynolds Costa, no Notepad++
O uso é permitido; a comercialização, proibida.
*/

function ExcelParaSQL(obj){let _campos=new Array,contagem=function(e,o,t,r,a){void 0===a&&(a=0);let n="",s=0,c=0;for(let r=0;r<e.length;r++){let a=e.substring(r,r+1);o.indexOf(a)>-1&&s++,t.indexOf(a)>-1&&c++}return s!=c+a&&(n=r),n},like=function(e,o){let t=function(e,o){let t=!0;o=o.trim().replace("(","").replace(")","");try{-1==e.indexOf(o)&&(console.warn('O campo "'+o+'" não foi encontrado e foi removido'),t=!1)}catch(e){}return t},r=new Array;return e.forEach((e=>{if(e.indexOf("<")>-1||e.indexOf(">")>-1){let a=e.indexOf("<")>-1?"<":">",n=e.split(a);a=" "+a+" ";const s=function(e){const o=e.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);if(!o)return!1;const t=parseInt(o[1],10),r=parseInt(o[2],10)-1,a=parseInt(o[3],10),n=new Date(a,r,t);if(n.getFullYear()!==a||n.getMonth()!==r||n.getDate()!==t)return!1;let s=[a,r+1,t];for(let e=0;e<3;e++)s[e]=s[e].toString().padStart(2,"0");return"'"+s.join("-")+"'"}(n[1]);let c=n[1].replace("/","").replace("(","").replace(")",""),i=!1;c==parseFloat(c)||s?s&&(n[1]=s):(i=!0,console.error("Os operadores >, >=, < e <= só estão disponíveis para datas e valores numéricos")),!i&&t(o,n[0])&&r.push((n[0]+a+n[1]).replace("> /",">= ").replace("< /","<= "))}else if(e.indexOf("=")>-1){let a=e.split("=");t(o,a[0])&&r.push(parseInt(a[1])==a[1]?a[0]+"="+a[1]:a[0]+" LIKE '%"+a[1].split(" ").join("%")+"%'")}else r.push(e)})),r},eColuna=function(e){if(!e||"string"!=typeof e)return!1;if(e=e.trim(),!/^[a-zA-Z_$]/.test(e))return!1;for(let o=1;o<e.length;o++)if(!/^[a-zA-Z0-9_$]/.test(e[o]))return!1;return!0};const todos=function(e,o){for(e=e.toString();e.indexOf("'")>-1;)e=e.replace("'","");let t=new Array;return o.forEach((o=>{t.push(o+" LIKE '%"+e+"%'")})),[t.join(" OR "),o]};if(this.e=function(e,o){return void 0===o&&(o=e,e=null),"("+like(o,e).join(" AND ")+")"},this.ou=function(e,o){return void 0===o&&(o=e,e=null),"("+like(o,e).join(" OR ")+")"},this.nao=function(e,o){return void 0===o&&(o=e,e=null),"NOT("+like([o],e)[0]+")"},void 0!==obj){let separador=",",erro="",retorno=!1,_campos_ret=new Array;if("object"==typeof obj){let e="",o=!1;for(x in obj)-1==["conteudo","separador","campos"].indexOf(x)&&(e='A chave "'+x+'" não é conhecida');e&&console.warn(e),"string"!=typeof obj.conteudo&&(erro='A chave "conteúdo" deve ser uma string'),(-1==["undefined","string"].indexOf(typeof obj.separador)||"string"==typeof obj.separador&&1!=obj.separador.length)&&(o=!0),erro||(o?erro='A chave "separador", se declarada, deve ser uma string de tamanho 1':void 0!==obj.separador&&(separador=obj.separador)),erro||(-1==["undefined","object"].indexOf(typeof obj.campos)?o=!0:"object"==typeof obj.campos&&obj.campos.forEach((e=>{"string"==typeof e&&eColuna(e)||(o=!0)})),o?erro='A chave "campos", se declarada, deve ser um array de colunas SQL':void 0!==obj.campos&&(_campos=obj.campos))}else erro="Parâmetro não declarado corretamente";if(!erro){var texto=obj.conteudo.toLowerCase();if(","!=separador)for(;texto.indexOf(separador)>-1;)texto=texto.replace(separador,",");const e=texto.indexOf("e(")>-1||texto.indexOf("ou(")>-1||texto.indexOf("nao(")>-1;var pseudocodigo=e||texto.indexOf(",")>-1||texto.indexOf("<")>-1||texto.indexOf("=")>-1||texto.indexOf(">")>-1;for(!e&&pseudocodigo&&(texto="e("+texto+")");texto.indexOf(">=")>-1;)texto=texto.replace(">=",">/");for(;texto.indexOf("<=")>-1;)texto=texto.replace("<=","</");e&&(erro=contagem(texto,"(",")","Número desigual de parênteses"),erro||(erro=contagem(texto,"<=>",",","Termos não declarados corretamente",1)))}if(!erro){texto=texto.replace(/\(/g,"(['").replace(/\)/g,"'])").replace(/,/g,"','");const substituir={",'(['":",","'ou([":"ou(['","'e([":"e(['","'nao([":"nao(['","])'":"])","'ou":"ou","'e":"e","'nao":"nao","''":"'"," ":""};for(x in substituir)for(;texto.indexOf(x)>-1;)texto=texto.replace(x,substituir[x]);texto=texto.replace(/e\(/g,"new ExcelParaSQL().e(").replace(/ou\(/g,"new ExcelParaSQL().ou(").replace(/nao\(/g,"new ExcelParaSQL().nao(");try{let texto_formatado=eval(texto),poss_col=new Array,poss_col2=new Array;if(["<","=",">"].forEach((e=>{texto.split(e).forEach((e=>{(e=e.split("[")).forEach((e=>{poss_col.push(e)}))}))})),poss_col.forEach((e=>{if(e.indexOf("'")>-1){let o=e.split("'");o=o[o.length-1],eColuna(o)&&-1==poss_col2.indexOf(o)&&(_campos.indexOf(o)>-1||!_campos.length)&&poss_col2.push(o)}})),texto=eval(texto.replace(/e\(/g,"e(['"+poss_col2.join("','")+"'],").replace(/ou\(/g,"ou(['"+poss_col2.join("','")+"'],").replace(/nao\(/g,"nao(['"+poss_col2.join("','")+"'],")),poss_col2.forEach((e=>{texto.indexOf(e)>-1&&_campos_ret.push(e)})),!pseudocodigo&&_campos.length){const e=todos(texto,_campos);texto=e[0],_campos_ret=e[1]}}catch(e){if(!pseudocodigo&&_campos.length){const e=todos(texto,_campos);texto=e[0],_campos_ret=e[1]}else erro="Erro desconhecido"}}if(erro)console.error(erro);else{for(texto=texto.toString();texto.indexOf("()")>-1;)texto=texto.replace("()","1");for(;texto.indexOf("  ")>-1;)texto=texto.replace("  "," ");retorno={sql:texto,campos:_campos_ret}}return retorno}}