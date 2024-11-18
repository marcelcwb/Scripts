import paramiko # type: ignore
import re
import ipaddress
from openpyxl import Workbook # type: ignore

# Função para se conectar ao switch e obter as VLANs, descrições, IPs e configurações adicionais
def get_vlans_and_configs_from_switch(host, username, password):
    # Configuração do cliente SSH
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    
    try:
        # Conectando ao switch via SSH
        ssh_client.connect(host, username=username, password=password)
        
        # Executando o comando 'show vlan brief' para obter informações sobre as VLANs
        stdin, stdout, stderr = ssh_client.exec_command('show vlan brief')
        vlan_output = stdout.read().decode()
        
        # Executando o comando 'show running-config' para obter as interfaces e suas configurações
        stdin, stdout, stderr = ssh_client.exec_command('show running-config')
        config_output = stdout.read().decode()
        
        # Fechando a conexão SSH
        ssh_client.close()
        
        # Regex para extrair as VLANs e descrições
        vlan_pattern = re.compile(r'(\d+)\s+([\w-]+)\s+.*\s+(\S+)\s+([\w\s]+)')
        
        # Regex para capturar os endereços IP e máscaras de subrede das interfaces VLAN
        ip_pattern = re.compile(r'interface Vlan(\d+)\s+.*?ip address (\S+)\s+(\S+)')
        
        # Lista para armazenar as VLANs, descrições e configurações
        vlans = []
        
        # Procurar todas as VLANs e suas descrições
        for match in vlan_pattern.finditer(vlan_output):
            vlan_id = match.group(1)
            vlan_name = match.group(4).strip()
            ip_address = None
            subnet_mask = None
            network_address = None
            
            # Procurar o IP e a máscara correspondente à VLAN
            for ip_match in ip_pattern.finditer(config_output):
                if ip_match.group(1) == vlan_id:
                    ip_address = ip_match.group(2)  # O IP é o segundo grupo da expressão regular
                    subnet_mask = ip_match.group(3)  # A máscara de subrede é o terceiro grupo
                    break
            
            # Calcular o barramento de rede (network address)
            if ip_address and subnet_mask:
                try:
                    # Usando a biblioteca ipaddress para calcular o barramento de rede
                    network = ipaddress.IPv4Network(f'{ip_address}/{subnet_mask}', strict=False)
                    network_address = network.network_address
                except ValueError:
                    network_address = "N/A"
            
            # Adicionar as informações da VLAN à lista
            vlans.append((vlan_id, vlan_name, ip_address if ip_address else "N/A", 
                          subnet_mask if subnet_mask else "N/A", network_address))
        
        return vlans
    except Exception as e:
        print(f"Erro ao se conectar ao switch: {e}")
        return []

# Função para criar a planilha Excel com as VLANs, descrições, IPs e configurações adicionais
def save_vlans_to_excel(vlans, filename="vlan_details_with_network_info.xlsx"):
    # Cria uma nova planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "VLANs"
    
    # Definindo o cabeçalho
    ws.append(["VLAN ID", "Descrição", "Endereço IP", "Máscara de Subrede", "Barramento de Rede"])
    
    # Adicionando as VLANs à planilha
    for vlan in vlans:
        ws.append([vlan[0], vlan[1], vlan[2], vlan[3], vlan[4]])
    
    # Salvando a planilha
    wb.save(filename)
    print(f"Planilha salva como {filename}")

# Função principal
def main():
    # Informações de conexão SSH
    host = input("Digite o IP ou hostname do switch Cisco: ")
    username = input("Digite o nome de usuário: ")
    password = input("Digite a senha: ")
    
    # Obter as VLANs e suas configurações do switch
    vlans = get_vlans_and_configs_from_switch(host, username, password)
    
    if vlans:
        print(f"Encontradas {len(vlans)} VLANs. Salvando na planilha...")
        save_vlans_to_excel(vlans)
    else:
        print("Nenhuma VLAN encontrada ou erro na conexão.")

if __name__ == "__main__":
    main()
