import os
import xmltodict
import pandas as pd

def list_xml_files(directory):
    files = os.listdir(directory)
    xml_files = [f for f in files if f.endswith('.xml')]
    return xml_files

def extract_data_from_nfe(xml_dict):
    nfe_proc = xml_dict.get('nfeProc')
    if not nfe_proc:
        return []
    
    nfe = nfe_proc.get('NFe')
    if not nfe:
        return []
    
    inf_nfe = nfe.get('infNFe', {})
    ide = inf_nfe.get('ide', {})
    emit = inf_nfe.get('emit', {})
    dest = inf_nfe.get('dest', {})
    det = inf_nfe.get('det', [])
    
    if not isinstance(det, list):
        det = [det]

    data_list = []
    for item in det:
        prod = item.get('prod', {})
        
        data = {
            'nNF': ide.get('nNF', ''),
            'dhEmi': ide.get('dhEmi', ''),
            'CNPJ_Emitente': emit.get('CNPJ', ''),
            'xFant_Emitente': emit.get('xFant', ''),
            'CNAE_Emitente': emit.get('CNAE', ''),
            'IE_Emitente': emit.get('IE', ''),
            'CNPJ_Destinatario': dest.get('CNPJ', ''),
            'xNome_Destinatario': dest.get('xNome', ''),
            'nItem': item.get('@nItem', ''),
            'Produto': prod.get('xProd', ''),
            'CFOP': prod.get('CFOP', ''),
            'NCM': prod.get('NCM', ''),
            'vProd': prod.get('vProd', '')
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
            with open(file_path, 'r', encoding='utf-8') as file:
                xml_content = file.read()
                xml_dict = xmltodict.parse(xml_content)
                data = extract_data_from_nfe(xml_dict)
                all_data.extend(data)
        except Exception as e:
            print(f"Erro ao processar o arquivo {xml_file}: {e}")
    
    if not all_data:
        print("Nenhuma nota fiscal válida encontrada.")
        return
    
    df = pd.DataFrame(all_data)
    output_file = os.path.join(directory, 'notas_fiscais_nfce.xlsx')
    df.to_excel(output_file, index=False)
    print(f"Planilha gerada com sucesso: {output_file}")

if __name__ == "__main__":
    main()
