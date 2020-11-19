from funcoes376 import automacao376 as aut
files = aut.import_txt()
trabalho, criticas = aut.valida_criticas(files)
resultado = aut.outputs(trabalho, criticas)