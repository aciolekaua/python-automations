import os
import xmltodict
import pandas as pd

def list_xml_files(directory):
    files = os.listdir(directory)
    xml_files = [f for f in files if f.endswith('.xml')]
    return xml_files

def extract_data_from_xml(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        xml_content = file.read()
        xml_dict = xmltodict.parse(xml_content)
        
        consultar_nfse_resposta = xml_dict.get('ConsultarNfseResposta')
        if not consultar_nfse_resposta:
            return []
        
        lista_nfse = consultar_nfse_resposta.get('ListaNfse')
        if not lista_nfse:
            return []
        
        comp_nfse = lista_nfse.get('CompNfse')
        if not comp_nfse:
            return []
        
        if not isinstance(comp_nfse, list):
            comp_nfse = [comp_nfse]
        
        data_list = []
        for comp in comp_nfse:
            nfse = comp.get('Nfse')
            if not nfse:
                continue
            
            inf_nfse = nfse.get('InfNfse', {})
            servico = inf_nfse.get('Servico', {}).get('Valores', {})
            prestador = inf_nfse.get('PrestadorServico', {}).get('IdentificacaoPrestador', {})
            tomador = inf_nfse.get('TomadorServico', {}).get('IdentificacaoTomador', {}).get('CpfCnpj', {})
            
            if tomador is None:
                cnpj_cpf_value = '0'
            else:
                cnpj = tomador.get('Cnpj', '')
                cpf = tomador.get('Cpf', '')
                cnpj_cpf_value = cnpj if cnpj else cpf
            
            discriminacao = inf_nfse.get('Servico', {}).get('Discriminacao', '')
            item_lista_servico = inf_nfse.get('Servico', {}).get('ItemListaServico', '')
            codigo_cnae = inf_nfse.get('Servico', {}).get('CodigoCnae', '')
                
            data = {
                'id': inf_nfse.get('@id', ''),
                'Numero': inf_nfse.get('Numero', ''),
                'CnpjPrestador': prestador.get('Cnpj', ''),
                'NomeFantasiaPrestador': inf_nfse.get('PrestadorServico', {}).get('NomeFantasia', ''),
                'ValorServicos': servico.get('ValorServicos', ''),
                'IssRetido': servico.get('IssRetido', ''),
                'ValorIss': servico.get('ValorIss', ''),
                'ValorIssRetido': servico.get('ValorIssRetido', ''),
                'Aliquota': servico.get('Aliquota', ''),
                'CpfCnpjTomador': cnpj_cpf_value,
                'ItemListaServico': item_lista_servico,
                'CodigoCnae': codigo_cnae,
                'Discriminacao': discriminacao,
            }
            
            data_list.append(data)
        
        return data_list

def main():
    directory = input("Digite o caminho do diretório onde os arquivos XML estão localizados: ")
    xml_files = list_xml_files(directory)
    
    if not xml_files:
        print("Nenhum arquivo XML encontrado no diretório especificado.")
        return
    
    all_data = []
    for xml_file in xml_files:
        file_path = os.path.join(directory, xml_file)
        try:
            data = extract_data_from_xml(file_path)
            all_data.extend(data)
        except Exception as e:
            print(f"Erro ao processar o arquivo {xml_file}: {e}")
    
    if not all_data:
        print("Nenhuma nota fiscal válida encontrada.")
        return
    
    df = pd.DataFrame(all_data)
    output_file = os.path.join(directory, 'notas_fiscais.xlsx')
    df.to_excel(output_file, index=False)
    print(f"Planilha gerada com sucesso: {output_file}")

if __name__ == "__main__":
    main()
