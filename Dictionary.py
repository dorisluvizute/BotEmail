def dictionary():
    dict = {
        "=41": "A",
        "=5A": "Z",
        "=61": "a",
        "=7A": "z",
        "=C0": "À",
        "=C1": "Á",
        "=C2": "Â",
        "=C3": "Ã",
        "=C4": "Ä",
        "=C5": "Å",
        "=C7": "Ç",
        "=C8": "È",
        "=C9": "É",
        "=CA": "Ê",
        "=CB": "Ë",
        "=CC": "Ì",
        "=CD": "Í",
        "=CE": "Î",
        "=CF": "Ï",
        "=D1": "Ñ",
        "=D2": "Ò",
        "=D3": "Ó",
        "=D4": "Ô",
        "=D5": "Õ",
        "=D6": "Ö",
        "=D9": "Ù",
        "=DA": "Ú",
        "=DB": "Û",
        "=DC": "Ü",
        "=DD": "Ý",
        "=E0": "à",
        "=E1": "á",
        "=E2": "â",
        "=E3": "ã",
        "=E4": "ä",
        "=E7": "ç",
        "=E8": "è",
        "=E9": "é",
        "=EA": "ê",
        "=EB": "ë",
        "=EC": "ì",
        "=ED": "í",
        "=EE": "î",
        "=EF": "ï",
        "=F1": "ñ",
        "=F2": "ò",
        "=F3": "ó",
        "=F4": "ô",
        "=F5": "õ",
        "=F6": "ö",
        "=F9": "ù",
        "=FA": "ú",
        "=FB": "û",
        "=FC": "ü",
        "=FD": "ý",
        "=FF": "ÿ"
    }

    return dict

def convert_ISO8859(name, dict):
    for key in dict.keys():
        name =name.replace(key, dict[key])
    
    return name

def body_text():
    text = ("Olá, seja Bem-Vindo!!!\n"
    + "Parabéns! Você agora faz parte da Qintess, empresa com trajetória de sucesso e grandes realizações na área de Tecnologia da Informação que nos posiciona hoje como um dos principais players em Transformação de Negócios da América Latina.\n"
    + "Para você começar nessa jornada de crescimento, precisamos de alguns documentos (anexados) para seguirmos com a sua contratação. É muito importante que você garanta a entrega de todos os documentos listados.\n"
    + "Basta encaminhá-los para o nosso Departamento Pessoal no e-mail documentos@qintess.com no prazo de 48 horas.\n"
    + "Orientações para a realização do exame médico:\n"
    + "SE VOCÊ RESIDIR EM SÃO PAULO E REGIÃO:\n"
    + "Comparecer na Clínica Delta Saúde.\n"
    + "Endereço: Rua França Pinto, 899 – Vila Mariana – SP\n"
    + "Horário: segunda a sexta–feira das 08 às 11 hrs (atendimento por ordem de chegada)\n"
    + "SE VOCE RESIDIR EM OUTROS ESTADOS OU INTERIOR DE SÃO PAULO:\n"
    + "Aguardar nosso contato via e-mail para orientações e agendamento do exame juntamente com a guia.\n"
    + "IMPORTANTE!\n"
    + "Sua data de início será confirmada após o recebimento de todos os documentos obrigatórios para cadastro.")

    return text

