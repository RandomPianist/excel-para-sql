function ExcelParaSQL(obj) {
	let campos = new Array();
	
	let contagem = function(texto, char1, char2, _erro, adc) {
		if (adc === undefined) adc = 0;
		let erro = "";
		let contA = 0;
		let contB = 0;
		for (let i = 0; i < texto.length; i++) {
			switch(texto.substring(i, i + 1)) {
				case char1:
					contA++;
					break;
				case char2:
					contB++;
					break;
			}
		}
		if (contA != contB + adc) erro = _erro;
		return erro;
	}

	let like = function(arr) {
		let resultado = new Array();
		arr.forEach((termo) => {
			if (termo.indexOf("=") > -1) {
				let partes = termo.split("=");
				if (campos.indexOf(partes[0].trim()) > -1 || !campos.length) resultado.push(partes[0] + " LIKE '%" + partes[1] + "%'");
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
	
	this.e = function(arr) {
		return "(" + like(arr).join(" AND ") + ")";
	}
	
	this.ou = function(arr) {
		return "(" + like(arr).join(" OR ") + ")";
	}
	
	this.nao = function(txt) {
		return "NOT(" + like([txt])[0] + ")";
	}
	
	if (obj !== undefined) {
		let separador = ",";
		let erro = "";
		let retorno = false;
		
		if (typeof obj == "object") {
			let aviso = "";
			let aux = false;
			for (x in obj) {
				if (["conteudo", "separador", "campos"].indexOf(x) == -1) aviso = 'A chave "' + x + '" não é conhecida';
			}
			if (aviso) console.warn(aviso);
			if (typeof obj.conteudo != "string") erro = 'A chave "conteúdo" deve ser uma string';
			if (["undefined", "string"].indexOf(typeof obj.separador) == -1) aux = true;
			else if (typeof obj.separador == "string") {
				if (obj.separador.length != 1) aux = true;
			}
			if (!erro) {
				if (aux) erro = 'A chave "separador", se declarada, deve ser uma string de tamanho 1';
				else separador = obj.separador;
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
				else campos = obj.campos;
			}
		} else erro = "Parâmetro não declarado corretamente";
		
		if (!erro) {
			var texto = obj.conteudo.toLowerCase();
			
			if (separador != ",") {
				while (texto.indexOf(separador) > -1) texto = texto.replace(separador, ",");
			}
			
			var funcoes = (
				texto.indexOf("e(") > -1 ||
				texto.indexOf("ou(") > -1 ||
				texto.indexOf("nao(") > -1
			);
			
			var pseudocodigo = (
				funcoes ||
				texto.indexOf(",") > -1 ||
				texto.indexOf("=") > -1
			);
			
			if (funcoes) {
				erro = contagem(texto, "(", ")", "Número desigual de parênteses");
				if (!erro) erro = contagem(texto, "=", ",", "Termos não declarados corretamente", 1);
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
				" " : ""
			};
			for (x in substituir) {
				while (texto.indexOf(x) > -1) texto = texto.replace(x, substituir[x]);
			}
			texto = texto.replace(/e\(/g,  "new ExcelParaSQL().e(")
						.replace(/ou\(/g,  "new ExcelParaSQL().ou(")
						.replace(/nao\(/g, "new ExcelParaSQL().nao(");
			try {
				texto = eval(texto);
				while (texto.indexOf("()") > -1) texto = texto.replace("()", "0");
			} catch(err) {
				if (!pseudocodigo && campos.length) {
					while (texto.indexOf("'") > -1) texto = texto.replace("'", "");
					let resultado = new Array();
					campos.forEach((campo) => {
						resultado.push(campo + " LIKE '%" + texto + "%'");
					});
					texto = resultado.join(" OR ");
				} else erro = "Erro desconhecido";
			}
		}
		if (erro) console.error(erro);
		else retorno = texto;
		return retorno;
	}
}