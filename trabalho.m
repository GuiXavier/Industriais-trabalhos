
% Script para gerar tabela de materiais da entrada de serviço conforme NTC 903100 (COPEL)

clear; clc;

% ENTRADA DE DADOS
disp('CADASTRO DE DADOS PARA DIMENSIONAMENTO DA ENTRADA DE SERVIÇO');
tensao = input('Tensão nominal [13.8 ou 34.5] kV: ');
tipoRamal = input('Tipo de Ramal [1=Aéreo, 2=Subterrâneo]: ');
tipoSubestacao = input('Tipo de Subestação [1=Ao tempo, 2=Cabina Alvenaria, 3=Cabina Metálica]: ');
ambiente = input('Ambiente [1=Normal, 2=Agressivo ou Litoral]: ');
potencia = input('Potência do transformador (kVA): ');

% INICIALIZAÇÃO DE TABELA DE MATERIAIS
materiais = {};
quantidade = {};

% SELEÇÃO BÁSICA DE MATERIAIS (pode ser expandido)
materiais{end+1} = 'Transformador de distribuição';
quantidade{end+1} = 1;

if tensao == 13.8
    materiais{end+1} = 'Chave Fusível 15 kV';
    quantidade{end+1} = 3;
    materiais{end+1} = 'Isolador Pilar 15 kV';
    quantidade{end+1} = 3;
    materiais{end+1} = 'Pára-Raios 15 kV';
    quantidade{end+1} = 3;
elseif tensao == 34.5
    materiais{end+1} = 'Chave Fusível 27 kV';
    quantidade{end+1} = 3;
    materiais{end+1} = 'Isolador Pilar 36 kV';
    quantidade{end+1} = 3;
    materiais{end+1} = 'Pára-Raios 27 kV';
    quantidade{end+1} = 3;
end

if tipoRamal == 1
    materiais{end+1} = 'Cabo coberto 35 mm²';
    quantidade{end+1} = 30; % metros
    materiais{end+1} = 'Poste de Concreto Duplo T';
    quantidade{end+1} = 1;
else
    materiais{end+1} = 'Eletroduto PVC rígido Ø100 mm';
    quantidade{end+1} = 20;
    materiais{end+1} = 'Caixa de Passagem 800x800';
    quantidade{end+1} = 1;
end

% Materiais adicionais comuns
materiais{end+1} = 'Haste de Aterramento 3 m aço cobreado';
quantidade{end+1} = 2;
materiais{end+1} = 'Condutor de aterramento 35 mm² nu';
quantidade{end+1} = 20;

% TABELA FINAL
T = table(materiais', quantidade', 'VariableNames', {'Material', 'Quantidade'});
disp('Relação de Materiais:');
disp(T);

% Exportar para Excel
writetable(T, 'materiais_entrada_servico.xlsx');
disp('Tabela exportada para materiais_entrada_servico.xlsx');
