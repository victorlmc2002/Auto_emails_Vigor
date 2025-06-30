class Devedores:
    def __init__(self, endereco, nome_inquilino, email_inquilino, valor, aluguel, iptu, cond, cotas_extras, repasse, taxa, tarifa, pdfboleto, pdfcond, pdfrepasse):
        self._endereco = endereco
        self._nome_inquilino = nome_inquilino
        self._email_inquilino = email_inquilino
        self._valor = self.converter_para_float(valor)
        self._aluguel = self.converter_para_float(aluguel)
        self._iptu = self.converter_para_float(iptu)
        self._cond = self.converter_para_float(cond)
        self._repasse = self.converter_para_float(repasse)
        self._taxa = self.converter_para_float(taxa)
        self._tarifa = self.converter_para_float(tarifa)
        self._cotas_extras = cotas_extras
        self._pdfboleto = pdfboleto
        self._pdfcond = pdfcond
        self._pdfrepasse = pdfrepasse
    
    def converter_para_float(self, valor):
        try:
            valor_float = float(valor)
            return round(valor_float, 2)  # Arredonda para 2 casas decimais
        except (TypeError, ValueError):
            return valor  # Se não for possível converter, mantém o valor original
    
    def __str__(self):
        #return f'{self._nome_inquilino} | {self._valor} | {self._aluguel} | {self._iptu} | {self._cond} | {self._cotas_extras} | {self._repasse} | {self._tarifa} | {self._pdfboleto} | {self._pdfcond} | {self._pdfrepasse}'
        return f'{self._pdfboleto}'