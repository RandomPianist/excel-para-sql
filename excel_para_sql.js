/*
EXCEL PARA SQL © 2024
Desenvolvido por Reynolds Costa, no Notepad++
O uso é permitido; a comercialização, proibida.
*/

function ExcelParaSQL(obj) {
	let _campos = new Array();
    let _tipos = new Array();
	
	let contagem = function(texto, char1, char2, _erro, adc) {
		if (adc === undefined) adc = 0;
		let erro = "";
		let contA = 0;
		let contB = 0;
		for (let i = 0; i < texto.length; i++) {
			let ref = texto.substring(i, i + 1);
			if (char1.indexOf(ref) > -1) contA++;
			if (char2.indexOf(ref) > -1) contB++;
		}
		if (contA != contB + adc) erro = _erro;
		return erro;
	}

	let like = function(arr, campos) {
		let verificar = function(__campos, chave) {
			let ret = true;
			chave = chave.trim().replace("(", "").replace(")", "");
			try {
				if (__campos.indexOf(chave) == -1) {
					console.warn('O campo "' + chave + '" não foi encontrado e foi removido');
					ret = false;
				}
			} catch(err) {}
			return ret;
		}
		
		let eData = function(dateString) {
			const regex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
			const match = dateString.match(regex);
			if (!match) return false;
			const day = parseInt(match[1], 10);
			const month = parseInt(match[2], 10) - 1;
			const year = parseInt(match[3], 10);
			const date = new Date(year, month, day);
			if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) return false;
			let resultado = [year, month + 1, day];
			for (let i = 0; i < 3; i++) resultado[i] = resultado[i].toString().padStart(2, "0");
			return "'" + resultado.join("-") + "'";
		}
		
		let resultado = new Array();
		arr.forEach((termo) => {
            if (typeof termo == "object") termo = termo[0];
			if (termo.indexOf("<") > -1 || termo.indexOf(">") > -1) {
				let operador = termo.indexOf("<") > -1 ? "<" : ">";
				let partes = termo.split(operador);
				operador = " " + operador + " ";
				const data = eData(partes[1]);
				let aux = partes[1].replace("/", "").replace("(", "").replace(")", "");
				let erro = false;
				if (aux != parseFloat(aux) && !data) {
					erro = true;
					console.error("Os operadores >, >=, < e <= só estão disponíveis para datas e valores numéricos");
				} else if (data) partes[1] = data;
				if (!erro && verificar(campos, partes[0])) resultado.push((partes[0] + operador + partes[1]).replace("> /", ">= ").replace("< /", "<= "));
			} else if (termo.indexOf("=") > -1) {
				let partes = termo.split("=");
				if (verificar(campos, partes[0])) resultado.push(parseInt(partes[1]) == partes[1] ? partes[0] + "=" + partes[1] : partes[0] + " LIKE '%" + partes[1].split(" ").join("%") + "%'");
			} else resultado.push(termo);
		});
		return resultado;
	}
	
	let eColuna = function(name) {
		if (!name || typeof name !== "string") return false;
		name = name.trim();
		if (!/^[a-zA-Z_$]/.test(name)) return false;
		for (let i = 1; i < name.length; i++) {
			if (!/^[a-zA-Z0-9_$]/.test(name[i])) return false;
		}
		return true;
	}
	
	const todos = function(txt, arr_campos, arr_tipos) {
        if (txt) {
            txt = txt.toString();
            while (txt.indexOf("'") > -1) txt = txt.replace("'", "");
            let resultado = new Array();
            for (let i = 0; i < arr_campos.length; i++) {
                if ((!arr_tipos[i] && parseInt(txt) == txt) || arr_tipos[i]) resultado.push(arr_tipos[i] ? arr_campos[i] + " LIKE '%" + txt + "%'" : arr_campos[i] + "=" + txt);
            }
            return [resultado.join(" OR "), arr_campos];
        }
        return ["1", arr_campos];        
	}
	
	this.e = function(campos, arr) {
		if (arr === undefined) {
			arr = campos;
			campos = null;
		}
		return "(" + like(arr, campos).join(" AND ") + ")";
	}
	
	this.ou = function(campos, arr) {
		if (arr === undefined) {
			arr = campos;
			campos = null;
		}
		return "(" + like(arr, campos).join(" OR ") + ")";
	}
	
	this.nao = function(campos, txt) {
		if (txt === undefined) {
			txt = campos;
			campos = null;
		}
		return "NOT(" + like([txt], campos)[0] + ")";
	}
	
	if (obj !== undefined) {
		let separador = ",";
		let erro = "";
		let retorno = false;
		let _campos_ret = new Array();
		
		if (typeof obj == "object") {
			let aviso = "";
			let aux = false;
			for (x in obj) {
				if (["conteudo", "separador", "campos", "tipos"].indexOf(x) == -1) aviso = 'A chave "' + x + '" não é conhecida';
			}
			if (aviso) console.warn(aviso);
			if (typeof obj.conteudo != "string") erro = 'A chave "conteúdo" deve ser uma string';
			if (["undefined", "string"].indexOf(typeof obj.separador) == -1) aux = true;
			else if (typeof obj.separador == "string") {
				if (obj.separador.length != 1) aux = true;
			}
			if (!erro) {
				if (aux) erro = 'A chave "separador", se declarada, deve ser uma string de tamanho 1';
				else if (obj.separador !== undefined) separador = obj.separador;
			}
			if (!erro) {
				if (["undefined", "object"].indexOf(typeof obj.campos) == -1) aux = true;
				else if (typeof obj.campos == "object") {
					obj.campos.forEach((campo) => {
						if (typeof campo == "string") {
							if (!eColuna(campo)) aux = true;
						} else aux = true;
					});
				}
				if (aux) erro = 'A chave "campos", se declarada, deve ser um array de colunas SQL';
				else if (obj.campos !== undefined) _campos = obj.campos;
			}
            if (!erro) {
				if (["undefined", "object"].indexOf(typeof obj.tipos) == -1) aux = true;
				else if (typeof obj.tipos == "object") {
					obj.tipos.forEach((tipo) => {
						if (!typeof tipo == "boolean") aux = true;
					});
				}
				if (aux) erro = 'A chave "tipos", se declarada, deve ser um array de booleanos, verdadeiros quando o campo correspondente for uma string';
				else if (obj.tipos !== undefined) _tipos = obj.tipos;
			}
            if (_tipos.length != _campos.length) erro = 'Os arrays "tipos" e "campos" devem ter o mesmo tamanho, pois são correspondentes';
		} else erro = "Parâmetro não declarado corretamente";
		
		if (!erro) {
			var texto = obj.conteudo.toLowerCase();
			
			if (separador != ",") {
				while (texto.indexOf(separador) > -1) texto = texto.replace(separador, ",");
			}
			
			const funcoes = (
				texto.indexOf("e(") > -1 ||
				texto.indexOf("ou(") > -1 ||
				texto.indexOf("nao(") > -1
			);
			
			var pseudocodigo = (
				funcoes ||
				texto.indexOf(",") > -1 ||
				texto.indexOf("<") > -1 ||
				texto.indexOf("=") > -1 ||
				texto.indexOf(">") > -1
			);
			
			if (!funcoes && pseudocodigo) texto = "e(" + texto + ")";
			
			while (texto.indexOf(">=") > -1) texto = texto.replace(">=", ">/");
			while (texto.indexOf("<=") > -1) texto = texto.replace("<=", "</");
			
			if (funcoes) {
				erro = contagem(texto, "(", ")", "Número desigual de parênteses");
				if (!erro) erro = contagem(texto, "<=>", ",", "Termos não declarados corretamente", 1);
			}
		}
		
		if (!erro) {
			texto = texto.replace(/\(/g, "(['")
						.replace(/\)/g, "'])")
						.replace(/,/g, "','");
			const substituir = {
				",'(['" : ",",
				"'ou([" : "ou(['",
				"'e([" : "e(['",
				"'nao([" : "nao(['",
				"])'" : "])",
				"'ou" : "ou",
				"'e" : "e",
				"'nao" : "nao",
				"''" : "'",
				" " : "",
                "([ou" : "(['ou",
                "([e" : "(['e",
                "([nao" : "(['nao"
			};
			for (x in substituir) {
				while (texto.indexOf(x) > -1) texto = texto.replace(x, substituir[x]);
			}
			texto = texto.replace(/e\(/g,  "new ExcelParaSQL().e(")
						.replace(/ou\(/g,  "new ExcelParaSQL().ou(")
						.replace(/nao\(/g, "new ExcelParaSQL().nao(");
			
			try {
				let texto_formatado = eval(texto);
				let poss_col = new Array();
				let poss_col2 = new Array();
				["<", "=", ">"].forEach((operador) => {
					texto.split(operador).forEach((aux) => {
						aux = aux.split("[");
						aux.forEach((aux2) => {
							poss_col.push(aux2);
						});
					});
				});
				poss_col.forEach((coluna) => {
					if (coluna.indexOf("'") > -1) {
						let aux = coluna.split("'");
						aux = aux[aux.length - 1];
						if (eColuna(aux) && poss_col2.indexOf(aux) == -1 && (_campos.indexOf(aux) > -1 || !_campos.length)) poss_col2.push(aux);
					}
				});
				texto = eval(
					texto.replace(/e\(/g,  "e(['"   + poss_col2.join("','") + "'],")
						.replace(/ou\(/g,  "ou(['"  + poss_col2.join("','") + "'],")
						.replace(/nao\(/g, "nao(['" + poss_col2.join("','") + "'],")
				);
				poss_col2.forEach((coluna) => {
					if (texto.indexOf(coluna) > -1) _campos_ret.push(coluna);
				});
				if (!pseudocodigo && _campos.length) {
					const aux = todos(texto, _campos, _tipos);
					texto = aux[0];
					_campos_ret = aux[1];
				}
			} catch(err) {
				if (!pseudocodigo && _campos.length) {
					const aux = todos(texto, _campos, _tipos);
					texto = aux[0];
					_campos_ret = aux[1];
				}
				else erro = "Erro desconhecido";
			}
		}
		if (!erro) {
			texto = texto.toString();
			while (texto.indexOf("()") > -1) texto = texto.replace("()", "1");
			while (texto.indexOf("  ") > -1) texto = texto.replace("  ", " ");
			retorno = {
				sql : encodeURIComponent(texto),
				campos : _campos_ret
			};
		} else console.error(erro);
		return retorno;
	}
}