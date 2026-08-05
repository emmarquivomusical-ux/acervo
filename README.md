<! DOCTYPE html>
<html lang="pt-BR">
<cabeça>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, inicial-scale=1.0, user-scalable=sim" />
    <título>Acervo · Escola Municipal de Música SP</título>
    <link rel="icon" href="data:image/SVG+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='-10 -10 20 20'><line x1='-8' y1='-8' x2='-8' y2='8' stroke='%23c9a84c' stroke-width='1.5' stroke-linecap='round'/><line x1='-8' y1='-8' x2='8' y2='-8' stroke='%23c9a84c' stroke-width='1.5' stroke-linecap='round'/><line x1='-8' y1='8' x2='8' y2='8' stroke='%23c9a84c' stroke-width='1.5' stroke-linecap='round'/>< Linha x1='-4' Y1='4' X2='4' Y2='4' Curso='%23C9A84C' Largura de Traço='1,5' Linha de Traço='Round'/><Linha X1='4' Y1='-8' X2='4' Y2='8' Curso='%23C9A84C' Largura de Traço='1,5' Stroke-Capitão='Redonda'/></SVG>" Tipo="imagem/SVG+XML" />

    <!-- supabaseClient JS Client -->
    <script src="https://cdn.jsdelivr.net/npm/@supabaseClient/supabaseClient-js@2"></script>

    <Estilo>
        /* === RESETAR E BASE === */
        * { margem: 0; preenchimento (padding): 0; tamanho de caixa: caixa de borda; }
        Corpo {
            família de fontes: system-ui, -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            Contexto: #F5F1eb;
            Cor: #2A2A28;
            Enchimento (REM): 1REM;
            Altura mínima: 100VH;
            display: flex;
            alinhar itens: centro;
            justificar-conteúdo: centro;
        }
 . container { max-width: 1300px; largura: 100%; margem: 0 auto; }

        /* === TELA DE LOGIN === */
 . login-wrapper {
            display: flex;
            alinhar itens: centro;
            justificar-conteúdo: centro;
            Altura mínima: 80VH;
            largura: 100%;
        }
 . Caixa de login {
            Contexto: #FFF;
            Raio de fronteira: 16px;
            Acolchoamento: 2,5Rem 2Rem;
            Largura máxima: 400px;
            largura: 100%;
            Sombra-caixa: 0 20px 60px rgba(0,0,0,0,0,08);
            Borda: 1px sólido #e0dcd4;
            Alinhar texto: centro;
        }
 . caixa de login. logo-icon { font-size: 3rem; display: block; margem-bottom: 0,5rem; }
 . login-box h1 { font-size: 1.8rem; Peso de fonte: 700; Cor: #2A2A28; }
 . caixa de login p { cor: #999; margem-fundo: 1,5rem; Tamanho da fonte: 0,9rem; }
 . etiqueta da caixa de login { display: block; Alinhamento de texto: esquerda; peso de fonte: 500; tamanho da fonte: 0,8rem; Cor: #555; margem-topo: 1rem; }
 . Entrada na caixa de login {
            largura: 100%;
            Enchimento de Acolchoamento: 0,7Rem 0,8Rem;
            Borda: 1px sólido #e0dcd4;
            Raio de fronteira: 8px;
            Tamanho da fonte: 0,95rem;
            Margem-Topo: 0,2Rem;
            Contexto: #FFF;
            transição: fronteira 0.2s;
        }
 . entrada da caixa de login:foco {
            Resumo: nenhum;
            Cor da borda: #C9A84C;
            Sombra de caixa: 0  0 0 3px rgba(201,168,76,0,15);
        }
 . caixa de login. btn-login {
            largura: 100%;
            Acolchoamento: 0,7REM;
            Contexto: #C9A84C;
            Cor: #FFF;
            fronteira: nenhuma;
            Raio de fronteira: 8px;
            Tamanho da fonte: 1rem;
            Peso da fonte: 600;
            cursor: ponteiro;
            Margem-topo: 1,5REM;
            transição: fundo 0.2s;
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(firstSheet);
                const obrasNovas = [];
                json.forEach(row => {
                    const titulo = (row['Título'] || row['titulo'] || '').toString().trim();
                    if (!titulo) return;
                    const tombo = (row['Tombo'] || row['tombo'] || '').toString().trim();
                    const compositor = (row['Compositor'] || row['compositor'] || '').toString().trim();
                    const editora = (row['Editora'] || row['editora'] || '').toString().trim();
                    const pasta = (row['Pasta'] || row['pasta'] || '').toString().trim();
                    const grupo = (row['Grupo'] || row['grupo'] || '').toString().trim();
                    const status = (row['Status'] || row['status'] || 'publico').toString().trim().toLowerCase();
                    obrasNovas.push({
                        tombo,
                        titulo,
                        compositor,
                        editora,
                        pasta,
                        grupo,
                        status: status === 'protegido' ? 'protegido' : 'publico'
                    });
                });
                if (obrasNovas.length === 0) {
                    mostrarToast('Nenhuma obra válida encontrada.', 'error');
                    return;
                }
                // 🔥 AQUI: usando "supabase" (não "supabaseClient")
                const { error } = await supabase
                    .from('obras')
                    .insert(obrasNovas);
                if (error) throw error;
                mostrarToast(`${obrasNovas.length} obras importadas com sucesso!`, 'success');
                await carregarDados();
            } catch (err) {
                console.error('Erro ao importar:', err);
                mostrarToast('Erro ao importar: ' + err.message, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    };
    input.click();
}

    // ============================================================
    //  RELÓGIO
    // ============================================================
    function iniciarRelogio() {
        function atualizar() {
            const now = new Date();
            const data = now.toLocaleDateString('pt-BR');
            const hora = String(now.getHours()).padStart(2, '0');
            const min = String(now.getMinutes()).padStart(2, '0');
            const seg = String(now.getSeconds()).padStart(2, '0');
            const horaFormatada = `${hora}:${min}:${seg}`;
            document.getElementById('digitalClock').textContent = `🕒 ${data} - ${horaFormatada}`;
            document.getElementById('topClock').textContent = `${data} ${horaFormatada}`;
        }
        atualizar();
        setInterval(atualizar, 1000);
    }

    // ============================================================
    //  TOAST
    // ============================================================
    function mostrarToast(msg, tipo = 'success') {
        const toast = document.getElementById('toast');
        toast.textContent = msg;
        toast.className = 'toast show ' + tipo;
        clearTimeout(toast._timer);
        toast._timer = setTimeout(() => {
            toast.classList.remove('show');
        }, 3500);
    }

    // ============================================================
    //  UTILITÁRIOS
    // ============================================================
    function escapeHtml(str) {
        if (!str) return '';
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // ============================================================
    //  INICIALIZAÇÃO
    // ============================================================
    async function init() {
        if (!isLoggedIn()) {
            document.getElementById('loginScreen').style.display = 'flex';
            document.getElementById('mainApp').classList.add('hidden');
            return;
        }
        document.getElementById('loginScreen').style.display = 'none';
        document.getElementById('mainApp').classList.remove('hidden');
        await carregarDados();
        renderizar();
        iniciarRelogio();
        atualizarContadoresMaster();
        atualizarDatalists();
    }

    document.addEventListener('DOMContentLoaded', function() {
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', function(e) {
            e.preventDefault();
            const user = document.getElementById('loginUser').value.trim();
            const pass = document.getElementById('loginPass').value.trim();
            if (user === 'admin' && pass === 'admin123') {
                localStorage.setItem('acervo_logged', 'true');
                document.getElementById('loginError').style.display = 'none';
                showMainApp();
            } else {
                document.getElementById('loginError').style.display = 'block';
            }
        });
    }
});
</script>

</Corpo>
</html>
