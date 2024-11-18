import paramiko # type: ignore
import time

# Função para conectar ao Check Point e alterar o servidor NTP
def alterar_servidor_ntp(host, usuario, senha, novo_ntp, porta=9922):
    try:
        # Estabelecer a conexão SSH com o dispositivo Check Point na porta 9922
        print(f"Conectando ao Check Point em {host} na porta {porta}...")
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())  # Ignora a verificação de chave do host
        client.connect(host, username=usuario, password=senha, port=porta)

        # Executar o comando para alterar o servidor NTP
        print("Alterando servidor NTP...")
        comando = f"set ntp servers {novo_ntp}"
        stdin, stdout, stderr = client.exec_command(comando)
        
        # Esperar a execução do comando
        time.sleep(2)  # Espera de 2 segundos para garantir que o comando seja executado
        
        # Checando a saída do comando
        output = stdout.read().decode()
        error = stderr.read().decode()

        if output:
            print(f"Saída: {output}")
        if error:
            print(f"Erro: {error}")
        
        # Confirmar a alteração com um comando de verificação
        print("Verificando a configuração do servidor NTP...")
        comando_verificacao = "show ntp servers"
        stdin, stdout, stderr = client.exec_command(comando_verificacao)
        output_verificacao = stdout.read().decode()
        
        if novo_ntp in output_verificacao:
            print(f"Servidor NTP alterado para: {novo_ntp}")
        else:
            print("Falha ao alterar o servidor NTP.")
    
    except Exception as e:
        print(f"Erro ao conectar ou executar o comando: {e}")
    
    finally:
        # Fechar a conexão SSH
        client.close()
        print("Conexão fechada.")

# Coletar os dados de acesso do usuário
host = input("Digite o endereço IP do Check Point: ")  # Endereço IP do firewall Check Point
usuario = input("Digite o nome de usuário: ")  # Nome de usuário
senha = input("Digite a senha: ")  # Senha
novo_ntp = input("Digite o novo servidor NTP: ")  # Novo servidor NTP
porta = input("Digite a porta SSH (padrão 9922): ")  # Porta SSH (opcional)

# Se a porta não for especificada, usa o valor padrão (9922)
if not porta:
    porta = 9922
else:
    porta = int(porta)

# Alterar o servidor NTP
alterar_servidor_ntp(host, usuario, senha, novo_ntp, porta)
