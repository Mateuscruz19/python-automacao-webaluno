# Verificador de Faltas Canvas PUCPR

Este é um script Python que automatiza a verificação de faltas no Canvas da PUCPR.

## Requisitos

- Python 3.8 ou superior
- Chrome instalado
- Conta PUCPR

## Instalação

1. Clone este repositório
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

3. Configure suas credenciais:
   - Abra o arquivo `.env`
   - Substitua `seu_usuario_aqui` pelo seu usuário PUCPR
   - Substitua `sua_senha_aqui` pela sua senha

## Uso

Execute o script com:
```bash
python canvas_checker.py
```

## Funcionalidades

- Login automático no Canvas
- Verificação de faltas (em desenvolvimento)
- Cálculo de faltas restantes (em desenvolvimento)

## Observações

- Mantenha suas credenciais seguras e não compartilhe o arquivo `.env`
- O script pode ser executado com ou sem interface gráfica (descomente a linha `chrome_options.add_argument("--headless")` para executar sem interface) #   p y t h o n - a u t o m a c a o - w e b a l u n o  
 