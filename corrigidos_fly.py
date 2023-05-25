# import auto_fly_stable as af
# import auto_fly as af
import auto_fly_normal as af
# import auto_fly_all_text as af
# import auto_fly_all_binary as af
# import time

def main():
    
    data = af.gerar_data_referencia()
    options = af.user_options
    
    """if options['verificar_execucao_anterior'].lower() == 's':
        af.verificar_se_ja_foi_emitido()"""

    af.abrir_e_logar(options['user'], options['password'], options['organograma'])
    
    af.gerar_rel_processos(data)
    af.abrir_rel_e_extrair_processos(corrigidos=True)

    if options['emitir_etiquetas_corrigidas'].lower() == 's':
        af.emitir_etiquetas()

    if options['emitir_planilhas_corrigidas'].lower() == 's':
        af.emitir_comprovantes_de_confirmacao_isolados()

    if options['emitir_protocolos_corrigidos'].lower() == 's':
        af.emitir_comprovantes_de_abertura_isolados()
    
    af.demonstrar_processos_sem_andamento()
    af.demonstrar_processos_com_andamento_incomum()

    if options['abrir_relatorios_corrigidos'].lower() == 's':
        af.abrir_relatorios_emitidos()
    else:
        af.acessar_no_menu('Relatórios', 'Gerenciador de relatórios')
    
    af.esperar_usuario_e_sair()


if __name__ == '__main__':
    main()
