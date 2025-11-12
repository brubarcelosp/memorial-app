// Configuração dinâmica de campos baseado no tipo de memorial
document.addEventListener('DOMContentLoaded', function() {
    const tipoEmp = document.getElementById('tipo_emp');
    const camposCondominio = document.getElementById('condominio_fields');
    const camposResumo = document.getElementById('resumo_fields');
    const camposAne = document.getElementById('ane_fields');
    const camposCoord = document.getElementById('coord_fields');
    const secaoUpload = document.getElementById('upload_section');
    const botaoExcel = document.getElementById('btn_excel');
    const aneDrop = document.getElementById('ane_drop');
    const grupoAneLargura = document.getElementById('ane_largura_group');

    // Alternar área não edificante
    aneDrop.addEventListener('change', function() {
        grupoAneLargura.style.display = this.value === 'Sim' ? 'block' : 'none';
    });

    // Atualizar campos visíveis baseado no tipo
    function atualizarVisibilidadeCampos() {
        const tipo = tipoEmp.value;
        
        // Resetar
        camposCondominio.style.display = 'none';
        camposResumo.style.display = 'none';
        camposAne.style.display = 'none';
        camposCoord.style.display = 'none';
        secaoUpload.style.display = 'none';
        botaoExcel.style.display = 'none';

        if (tipo === 'condominio' || tipo === 'loteamento') {
            if (tipo === 'condominio') {
                camposCondominio.style.display = 'block';
            }
            camposAne.style.display = 'block';
            camposCoord.style.display = 'block';
            secaoUpload.style.display = 'block';
            botaoExcel.style.display = 'block';
        } else if (tipo === 'memorial_resumo') {
            camposResumo.style.display = 'block';
        } else if (tipo === 'solicitacao_analise') {
            camposResumo.style.display = 'block';
        } else if (tipo === 'unificacao' || tipo === 'desmembramento' || tipo === 'unif_desm') {
            camposCoord.style.display = 'block';
            secaoUpload.style.display = 'block';
            botaoExcel.style.display = 'block';
        }
    }

    tipoEmp.addEventListener('change', atualizarVisibilidadeCampos);
    atualizarVisibilidadeCampos();

    // Upload de arquivos
    const botaoUpload = document.getElementById('btn_upload');
    const uploadArquivo = document.getElementById('file_upload');
    const statusUpload = document.getElementById('upload_status');

    botaoUpload.addEventListener('click', async function() {
        const arquivos = uploadArquivo.files;
        if (arquivos.length === 0) {
            mostrarMensagem('Por favor, selecione pelo menos um arquivo', 'error');
            return;
        }

        const dadosFormulario = new FormData();
        for (let arquivo of arquivos) {
            dadosFormulario.append('files', arquivo);
        }

        botaoUpload.disabled = true;
        botaoUpload.innerHTML = '<span class="loading"></span> Enviando...';

        try {
            const resposta = await fetch('/api/upload', {
                method: 'POST',
                body: dadosFormulario
            });

            const dados = await resposta.json();
            
            if (dados.success) {
                statusUpload.innerHTML = `
                    <strong>✅ ${dados.count} arquivo(s) carregado(s) com sucesso!</strong><br>
                    Arquivos: ${dados.files.join(', ')}
                `;
                mostrarMensagem(`✅ ${dados.count} arquivo(s) anexado(s) com sucesso!`, 'success');
            } else {
                mostrarMensagem('Erro ao fazer upload: ' + (dados.error || 'Erro desconhecido'), 'error');
            }
        } catch (erro) {
            mostrarMensagem('Erro ao fazer upload: ' + erro.message, 'error');
        } finally {
            botaoUpload.disabled = false;
            botaoUpload.innerHTML = '📎 Anexar Arquivos';
        }
    });

    // Geração de documento
    const formulario = document.getElementById('memorialForm');
    const botaoGerar = document.getElementById('btn_gerar');

    formulario.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        const dadosFormulario = new FormData(formulario);
        const dados = {};
        
        // Converter FormData para objeto
        for (let [chave, valor] of dadosFormulario.entries()) {
            if (dados[chave]) {
                // Se já existe, transformar em array
                if (Array.isArray(dados[chave])) {
                    dados[chave].push(valor);
                } else {
                    dados[chave] = [dados[chave], valor];
                }
            } else {
                dados[chave] = valor;
            }
        }

        // Tratar checkboxes
        dados.has_ai = document.getElementById('has_ai').checked;
        dados.has_restricao = document.getElementById('has_restricao').checked;

        // Tratar múltiplos usos
        const selecaoUsos = document.getElementById('usos_multi');
        if (selecaoUsos) {
            dados.usos_multi = Array.from(selecaoUsos.selectedOptions).map(opt => opt.value);
        }

        botaoGerar.disabled = true;
        botaoGerar.innerHTML = '<span class="loading"></span> Gerando...';

        try {
            const resposta = await fetch('/api/generate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(dados)
            });

            const resultado = await resposta.json();
            
            if (resultado.success) {
                mostrarMensagem('✅ Documento gerado com sucesso!', 'success');
                
                // Fazer download usando fetch com blob (mais confiável)
                try {
                    const respostaDownload = await fetch(resultado.download_url);
                    if (!respostaDownload.ok) {
                        throw new Error('Erro ao baixar arquivo');
                    }
                    
                    const blob = await respostaDownload.blob();
                    const url = window.URL.createObjectURL(blob);
                    const linkDownload = document.createElement('a');
                    linkDownload.href = url;
                    linkDownload.download = resultado.filename;
                    document.body.appendChild(linkDownload);
                    linkDownload.click();
                    document.body.removeChild(linkDownload);
                    window.URL.revokeObjectURL(url);
                } catch (erroDownload) {
                    console.error('Erro no download:', erroDownload);
                    // Fallback: tentar método antigo
                    const linkDownload = document.createElement('a');
                    linkDownload.href = resultado.download_url;
                    linkDownload.download = resultado.filename;
                    linkDownload.target = '_blank';
                    linkDownload.click();
                }
            } else {
                mostrarMensagem('❌ Erro ao gerar documento: ' + (resultado.error || 'Erro desconhecido'), 'error');
                if (resultado.traceback) {
                    console.error(resultado.traceback);
                }
            }
        } catch (erro) {
            mostrarMensagem('❌ Erro ao gerar documento: ' + erro.message, 'error');
        } finally {
            botaoGerar.disabled = false;
            botaoGerar.innerHTML = '📄 Gerar DOCX';
        }
    });

    // Geração de Excel
    botaoExcel.addEventListener('click', async function() {
        const dadosFormulario = new FormData(formulario);
        const dados = {};
        
        for (let [chave, valor] of dadosFormulario.entries()) {
            dados[chave] = valor;
        }

        dados.has_ai = document.getElementById('has_ai').checked;
        dados.has_restricao = document.getElementById('has_restricao').checked;

        const selecaoUsos = document.getElementById('usos_multi');
        if (selecaoUsos) {
            dados.usos_multi = Array.from(selecaoUsos.selectedOptions).map(opt => opt.value);
        }

        botaoExcel.disabled = true;
        botaoExcel.innerHTML = '<span class="loading"></span> Gerando...';

        try {
            const resposta = await fetch('/api/generate-excel', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(dados)
            });

            const resultado = await resposta.json();
            
            if (resultado.success) {
                mostrarMensagem('✅ Planilha gerada com sucesso!', 'success');
                const linkDownload = document.createElement('a');
                linkDownload.href = resultado.download_url;
                linkDownload.download = resultado.filename;
                linkDownload.click();
            } else {
                mostrarMensagem('❌ Erro ao gerar planilha: ' + (resultado.error || 'Erro desconhecido'), 'error');
            }
        } catch (erro) {
            mostrarMensagem('❌ Erro ao gerar planilha: ' + erro.message, 'error');
        } finally {
            botaoExcel.disabled = false;
            botaoExcel.innerHTML = '📊 Baixar Excel';
        }
    });
});

function mostrarMensagem(mensagem, tipo) {
    const divMensagens = document.getElementById('messages');
    const divMensagem = document.createElement('div');
    divMensagem.className = `message ${tipo}`;
    divMensagem.textContent = mensagem;
    divMensagens.appendChild(divMensagem);
    
    // Remover após 5 segundos
    setTimeout(() => {
        divMensagem.remove();
    }, 5000);
}

