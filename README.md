# consultahoraspontomais
Script para automatizar a consulta do registro de horas no site pontomais

#### Para executar via docker é necessário preenchimento dos parametros vazios:
docker run --privileged -e LOG_LEVEL='WARNING' -e LOGIN_PONTOMAIS='' -e SENHA_PONTOMAIS='' -e LOGIN_EMAIL='' -e SENHA_EMAIL='' -e SMPT_SERVER='' -e INTERVALO_VERIFICACAO=10 -e ASSUNTO_EMAIL='PontoMais' -e DRIVER_PADRAO='Firefox' -p 4000:4000 -d -it marcelogoncalvesdocker/consultahoraspontomais 
