using System;
using System.Net.Http;
using System.Net;
using System.Web.Http;
using Primavera.WebAPI.Integration;
using StpBE100;
using System.Collections.Generic;
using System.Globalization;
using StdBE100;
using static StpBE100.StpBETipos;
using System.Linq;
using System.Diagnostics.Contracts;
using BasBE100;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using StdPlatBS100;
using System.Collections;
using System.Web.Http.Description;
using System.Security.Cryptography;

namespace ADWebExtensibilidade
{
    public class OficioDto
    {
        public string Codigo { get; set; }
        public string Assunto { get; set; }
        public string Data { get; set; }
        public string Remetente { get; set; }
        public string Email { get; set; }
        public string Texto1 { get; set; }
        public string Texto2 { get; set; }
        public string Template { get; set; }
        public string Createdby { get; set; }
        public string Texto3 { get; set; }
        public string Obra { get; set; }
        public string Updatedby { get; set; }
        public string Isactive { get; set; }
        public string donoObra { get; set; }
        public string Morada { get; set; }
        public string Localidade { get; set; }
        public string CodPostal { get; set; }
        public string CodPostalLocal { get; set; }
        public string Anexos { get; set; }
        public string Texto4 { get; set; }
        public string Texto5 { get; set; }
        public string Estado { get; set; }
        public string Atencao { get; set; }
    }
    public class CriarPedidoRequest
    {
        public string Cliente { get; set; }
        public string DescricaoObjecto { get; set; }
        public string DescricaoProblema { get; set; }
        public string Origem { get; set; }
        public string TipoProcesso { get; set; }
        public string Prioridade { get; set; }
        public string Tecnico { get; set; }
        public string ObjectoID { get; set; }
        public string TipoDoc { get; set; }
        public string Serie { get; set; }
        public string Estado { get; set; }
        public string Seccao { get; set; }
        public string ComoReproduzir { get; set; }
        public string Contacto { get; set; }

        public string ContratoID { get; set; }
        public string Codigo { get; set; }

        public DateTime datahoraabertura { get; set; }
        public DateTime datahorafimprevista { get; set; }
    }
    public class CriarIntervencaoRequest
    {
        public string ProcessoID { get; set; }
        public string TipoIntervencao { get; set; }
        public int Duracao { get; set; }
        public int DuracaoReal { get; set; }
        public DateTime DataHoraInicio { get; set; }
        public DateTime DataHoraFim { get; set; }
        public string Tecnico { get; set; }
        public string EstadoAnt { get; set; }
        public string Estado { get; set; }
        public string SeccaoAnt { get; set; }
        public string Seccao { get; set; }
        public string Utilizador { get; set; }
        public string DescricaoResposta { get; set; }
        public List<dynamic> Artigos { get; set; } // Lista de artigos
    }

    public class ParteDiariaDto
    {
        // Campos para COP_FichasPessoal
        public string ID { get; set; }
        public int Numero { get; set; }
        public string ObraID { get; set; }
        public DateTime Data { get; set; }
        public string Encarregado { get; set; }
        public string Notas { get; set; }
        public string CabecMovCBLID { get; set; }
        public int LigaCBL { get; set; }
        public string CriadoPor { get; set; }
        public string Utilizador { get; set; }
        public DateTime DataUltimaActualizacao { get; set; }
        public string DocumentoID { get; set; }
        public string TipoEntidade { get; set; }
        public string SubEmpreiteiroID { get; set; }
        public int? ColaboradorID { get; set; }
        public int Validado { get; set; }

        // Campos para COP_FichasPessoalItems
        public string FichasPessoalID { get; set; }
        public string ComponenteID { get; set; }
        public string Funcionario { get; set; }
        public int ClasseID { get; set; }
        public string Fornecedor { get; set; }
        public int SubEmpID { get; set; }
        public int NumHoras { get; set; }
        public decimal PrecoUnit { get; set; }
        public string SEPessoalID { get; set; }
        public TimeSpan? ManhaInicio { get; set; }
        public TimeSpan? ManhaFim { get; set; }
        public TimeSpan? TardeInicio { get; set; }
        public TimeSpan? TardeFim { get; set; }
        public int TotalHoras { get; set; }
        public int Integrado { get; set; }
        public int? TipoHoraID { get; set; }
        public int? FuncaoID { get; set; }
        public string ItemId { get; set; }
    }
    public class EmailRequest
    {
        public string PedidoId { get; set; }
        public string UltimaNumIntervencao { get; set; }
    }


    public class AtualizaDataPedidoDto
    {
        public Guid Id { get; set; }
        public DateTime NovaData { get; set; }
    }

    [RoutePrefix("ServicosTecnicos")]
    public class ServicosTecnicosController : ApiController
    {

        [Authorize]
        [Route("LstUltimoPedido")]
        [HttpGet]
        public HttpResponseMessage LstUltimoPedido()
        {
            try
            {
                string query = $@"SELECT TOP 1 *
FROM STP_processos
WHERE ID IS NOT NULL
ORDER BY NumProcesso DESC;
";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum pedido encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter pedido: {ex.Message}");
            }
        }


        [Authorize]
        [Route("VerificaDataPedidoAtualiza")]
        [HttpPost]
        public HttpResponseMessage VerificaDataPedidoAtualiza([FromBody] AtualizaDataPedidoDto dto)
        {

            try
            {
                string query = $@" SELECT TOP 1 DataHoraAbertura
                                   FROM STP_processos
                                   WHERE ID = '{dto.Id}'
                                   ORDER BY NumProcesso DESC";


                var result = ProductContext.MotorLE.Consulta(query);

                if (result == null || result.NumLinhas() == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound,
                        "Nenhum processo encontrado com o ID informado.");
                }

                DateTime dataBanco = result.DaValor<DateTime>("DataHoraAbertura");


                // 👉 Só atualiza se a NovaData for menor
                if (dto.NovaData < dataBanco)
                {
                    string updateQuery = $@"
                    UPDATE STP_processos
                    SET DataHoraAbertura = DATEADD(MINUTE, -5, '{dto.NovaData:yyyy-MM-dd HH:mm:ss}')
                    WHERE ID = '{dto.Id}'";

                    int rows = ProductContext.MotorLE.DSO.ExecuteSQL(updateQuery);

                    return Request.CreateResponse(HttpStatusCode.OK,
                        $"Data atualizada com sucesso. Registros afetados: {rows}");
                }
                else
                {
                    return Request.CreateResponse(HttpStatusCode.OK,
                        "Nenhuma atualização necessária (a data recebida não é menor).");
                }
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter pedido: {ex.Message}");
            }
        }


        [Authorize]
        [Route("ObterPedidos")]
        [HttpGet]
        public HttpResponseMessage ObterPedidos()
        {
            try
            {
                string query = $@"SELECT  C.Nome, STP.*
FROM STP_processos AS STP
INNER JOIN Clientes AS C ON STP.Cliente = C.Cliente
where STP.Fechado != 1

";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum pedido encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter pedido: {ex.Message}");
            }
        }


        [Authorize]
        [Route("AreaclientListarpedidos/{IDCliente}")]
        [HttpGet]
        public HttpResponseMessage AreaclientListarpedidos(string IDCliente)
        {
            try
            {
                string query = $@"SELECT 
                                p.*, 
                                i.*, 
                                t.Nome AS NomeTecnico, 
                                e.Descricao AS DescricaoEstado
                            FROM 
                                STP_Processos p
                            LEFT JOIN 
                                STP_Intervencoes i ON i.ProcessoID = p.ID
                            LEFT JOIN 
                                STP_Tecnicos t ON i.Tecnico = t.Tecnico
                            LEFT JOIN 
                                STP_Estados e ON p.Estado = e.Estado
                            WHERE 
                                p.Cliente = '{IDCliente}'
                            ORDER BY 
                                p.NumProcesso DESC, 
                                i.Interv DESC;";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }




        [Authorize]
        [Route("ListarPedidosTecnico/{IDTecnico}")]
        [HttpGet]
        public HttpResponseMessage ListarPedidosTecnico(string IDTecnico)
        {
            try
            {
                string query = $@"SELECT  p.*, 
                                i.*, 
                                t.Nome AS NomeTecnico, 
                                e.Descricao AS DescricaoEstado
                            FROM STP_Processos p
                            LEFT JOIN 
                                STP_Intervencoes i ON i.ProcessoID = p.ID
                            LEFT JOIN 
                                STP_Tecnicos t ON i.Tecnico = t.Tecnico
                            LEFT JOIN 
                                STP_Estados e ON p.Estado = e.Estado
                            WHERE 
                                p.Tecnico = '{IDTecnico}'
                            ORDER BY 
                                p.NumProcesso DESC, 
                                i.Interv DESC;";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ObterInfoEmail/{PedidoId}/{UltimaNumIntervencao}")]
        [HttpGet]
        public HttpResponseMessage ObterInfoEmail(string PedidoId, string UltimaNumIntervencao)
        {
            try
            {
                var queryProcesso = $@"SELECT DescricaoObjecto,* FROM STP_Processos  Where ID = '{PedidoId}'";
                var Pedido = ProductContext.MotorLE.Consulta(queryProcesso);


                string dataHora = Pedido.DaValor<string>("DataHoraAbertura");
                string data = dataHora.Split(' ')[0];

                var queryIntervencao = $@"SELECT * FROM STP_Intervencoes where ProcessoID = '{PedidoId}' AND Interv = '{UltimaNumIntervencao}'";
                var intervencao = ProductContext.MotorLE.Consulta(queryIntervencao);

                var queryTecnico = $@"SELECT * FROM STP_Tecnicos where Tecnico = '{intervencao.DaValor<string>("Tecnico")}'";
                var tecnico = ProductContext.MotorLE.Consulta(queryTecnico);

                var queryEstado = $@"SELECT * FROM STP_Estados Where Estado = '{intervencao.DaValor<string>("Estado")}'";
                var estado = ProductContext.MotorLE.Consulta(queryEstado);

                string dataHora2 = intervencao.DaValor<string>("DataHoraInicio");
                string data2 = dataHora2.Split(' ')[0];

                var queryContacto = $@" SELECT PrimeiroNome,UltimoNome,* FROM Contactos where Contacto = '{Pedido.DaValor<string>("Contacto")}'";
                var contacto = ProductContext.MotorLE.Consulta(queryContacto);
                string nomeCompleto = $"{contacto.DaValor<string>("PrimeiroNome")}  {contacto.DaValor<string>("UltimoNome")}";

                var resposta = new
                {
                    Pedido = Pedido.DaValor<string>("DescricaoObjecto"),
                    Estado = estado.DaValor<string>("Descricao"),
                    Contacto = nomeCompleto,
                    HoraInicioPedido = data,
                    NumIntervencao = UltimaNumIntervencao,
                    Tecnico = tecnico.DaValor<string>("Nome"),
                    HoraInicioIntervencao = data2,
                    NumProcesso = Pedido.DaValor<string>("NumProcesso")
                };

                return Request.CreateResponse(HttpStatusCode.Created, resposta);
            }
            catch (Exception ex)
            {
                // Captura e retorna erros do servidor
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao adicionar o contacto: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ObterContactoIntervencao/{id}")]
        [HttpGet]
        public HttpResponseMessage ObterContactoPedido(string id)
        {
            try
            {
                // Verifica se o ID foi fornecido
                if (string.IsNullOrEmpty(id))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "ID não fornecido.");
                }

                // Consulta para obter o pedido
                var query2 = $@"SELECT Contacto,* FROM STP_Processos WHERE ID = '{id}'";
                var pedido = ProductContext.MotorLE.Consulta(query2);

                // Verifica se o pedido existe
                if (pedido.NumLinhas() == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Pedido não encontrado.");
                }

                // Verifica se o campo Contacto existe no pedido
                var contacto = pedido.DaValor<string>("Contacto");
                if (string.IsNullOrEmpty(contacto))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Contacto não encontrado no pedido.");
                }

                // Consulta para obter os contactos com o valor do campo Contacto
                var query3 = $@"SELECT Id,* FROM Contactos WHERE Contacto = '{contacto}'";
                var contact = ProductContext.MotorLE.Consulta(query3);

                // Verifica se algum contacto foi encontrado
                if (contact.NumLinhas() == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Contacto não encontrado.");
                }

                // Verifica se o campo Id do contacto existe
                var contactoId = contact.DaValor<string>("Id");
                if (string.IsNullOrEmpty(contactoId))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "ID do contacto não encontrado.");
                }

                // Consulta para obter as linhas de contacto
                var query4 = $@"SELECT * FROM LinhasContactoEntidades WHERE IDContacto = '{contactoId}'";
                var response = ProductContext.MotorLE.Consulta(query4);

                // Verifica se as linhas de contacto existem
                if (response.NumLinhas() > 0)
                {
                    return Request.CreateResponse(HttpStatusCode.OK, response);
                }
                else
                {
                    return Request.CreateResponse(HttpStatusCode.OK, "Não há informações disponíveis.");
                }
            }
            catch (Exception ex)
            {
                // Captura e retorna erros do servidor
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter os contactos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ObterInfoContrato/{id}")]
        [HttpGet]
        public HttpResponseMessage ObterInfoContrato(string id)
        {
            try
            {
                // Verifica se o ID foi fornecido
                if (string.IsNullOrEmpty(id))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "ID não fornecido.");
                }

                // Consulta para obter o pedido
                var query2 = $@"select pcm.*, stp.*
                                from PCM_Contratos pcm
                                join STP_Contratos stp on pcm.id = stp.id
                                where pcm.Entidade = '{id}';";
                var response = ProductContext.MotorLE.Consulta(query2);


                return Request.CreateResponse(HttpStatusCode.OK, response);

            }
            catch (Exception ex)
            {
                // Captura e retorna erros do servidor
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter os contactos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("EliminarPedido/{id}")]
        [HttpGet]
        public HttpResponseMessage EliminarPedido(string id)
        {
            try
            {
                ProductContext.MotorLE.ServicosTecnicos.Processos.RemoveID(id);

                return Request.CreateResponse(HttpStatusCode.OK, "Removido");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter estados: {ex.Message}");
            }
        }
        [Authorize]
        [Route("EliminarIntervencao/{id}")]
        [HttpGet]
        public HttpResponseMessage EliminarIntervencao(string id)
        {
            try
            {
                var interObjeto = ProductContext.MotorLE.ServicosTecnicos.Intervencoes.Edita(id);

                ProductContext.MotorLE.ServicosTecnicos.Intervencoes.Remove(interObjeto);

                return Request.CreateResponse(HttpStatusCode.OK, "Removido");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter estados: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetIntervencao/{id}")]
        [HttpGet]
        public HttpResponseMessage GetIntervencao(string id)
        {
            try
            {
                var response = ProductContext.MotorLE.ServicosTecnicos.Intervencoes.Edita(id);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum estado encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter estados: {ex.Message}");
            }
        }
        [Authorize]
        [Route("LstEstadosTodos")]
        [HttpGet]
        public HttpResponseMessage LstEstadosTodos()
        {
            try
            {
                var response = ProductContext.MotorLE.ServicosTecnicos.Estados.LstEstadosTodos();

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum estado encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter estados: {ex.Message}");
            }
        }

        [Authorize]
        [Route("LstTodosDocumentos")]
        [HttpGet]
        public HttpResponseMessage LstTodosDocumentos()
        {
            try
            {
                var response = ProductContext.MotorLE.ServicosTecnicos.DocumentosPedido.LstTodosDocumentos();

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum documento encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter documentos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("LstOrigensProcessos")]
        [HttpGet]
        public HttpResponseMessage LstOrigensProcessos()
        {
            try
            {
                var response = ProductContext.MotorLE.ServicosTecnicos.OrigensProcesso.LstOrigensProcessos();

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma origem de processo encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter origens de processos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("LstDocumentosProcesso")]
        [HttpGet]
        public HttpResponseMessage LstDocumentosProcesso()
        {
            try
            {
                var response = ProductContext.MotorLE.ServicosTecnicos.DocumentosProcesso.LstDocumentosProcesso();

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum documento encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter documentos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("LstTecnicosTodos")]
        [HttpGet]
        public HttpResponseMessage LstTecnicosTodos()
        {
            try
            {
                var response = ProductContext.MotorLE.ServicosTecnicos.Tecnicos.LstTecnicosTodos();

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum técnico encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter técnicos: {ex.Message}");
            }
        }
        [Authorize]
        [Route("LstTiposIntervencao")]
        [HttpGet]
        public HttpResponseMessage LstTiposIntervencao()
        {
            try
            {
                string query = "SELECT * FROM STP_TiposIntervencao";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum técnico encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter técnicos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetTempoDeslocacao/{processoID}")]
        [HttpGet]
        public HttpResponseMessage GetTempoDeslocacao(string processoID)
        {
            try
            {
                string query = $@"SELECT C.CDU_TempoDeslocacao FROM STP_Processos AS SP INNER JOIN Clientes AS C ON SP.Cliente = C.Cliente WHERE SP.ID = '{processoID}'";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tempo encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter tempos de Deslocacao: {ex.Message}");
            }
        }

        [Authorize]
        [Route("LstObjectos")]
        [HttpGet]
        public HttpResponseMessage LstObjectos()
        {
            try
            {
                var response = ProductContext.MotorLE.ServicosTecnicos.Objectos.LstObjectos();

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarObjetos")]
        [HttpGet]
        public HttpResponseMessage ListarObjetos()
        {
            try
            {
                string query = "SELECT * FROM STP_Objectos";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetObjetoID/{Objeto}")]
        [HttpGet]
        public HttpResponseMessage GetObjetoID(string Objeto)
        {
            try
            {
                string query = $@"SELECT ID FROM STP_Objectos where Objecto = '{Objeto}'";
                var response = ProductContext.MotorLE.Consulta(query);


                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarContactos/{IDCliente}")]
        [HttpGet]
        public HttpResponseMessage ListarContactos(string IDCliente)
        {
            try
            {
                string query = $@"SELECT
                                CO.Contacto,
                                CO.PrimeiroNome, 
                                CO.UltimoNome
                            FROM 
                                LinhasContactoEntidades AS L
                            INNER JOIN 
                                Clientes AS C
                                ON L.Entidade = C.Cliente
                            INNER JOIN 
                                Contactos AS CO
                                ON CO.Id = L.IDContacto
                            WHERE 
                                L.Entidade = '{IDCliente}';";

                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum contacto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter contactos: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarTiposProcesso")]
        [HttpGet]
        public HttpResponseMessage ListarTiposProcesso()
        {
            try
            {
                string query = "SELECT * FROM STP_TiposProcesso";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de processo encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter tipos de processo: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarTiposPrioridades")]
        [HttpGet]
        public HttpResponseMessage ListarTiposPrioridades()
        {
            try
            {
                string query = "SELECT * FROM STP_Prioridades";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma prioridade encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter tipos de prioridade: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarSeccoes")]
        [HttpGet]
        public HttpResponseMessage ListarSeccoes()
        {
            try
            {
                string query = "SELECT * FROM STP_Seccoes";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma seção encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter seções: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarContratos/{IDCliente}")]
        [HttpGet]
        public HttpResponseMessage ListarContratos(string IDCliente)
        {
            try
            {
                string query = $@"SELECT *
                        FROM 
                            Clientes AS C
                        INNER JOIN 
                            PCM_Contratos AS CO
                            ON CO.Entidade = C.Cliente
                        WHERE 
                            C.Cliente = '{IDCliente}'
                        AND estado != 5 AND estado !=6;";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma seção encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter seções: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarSeccaoUtilizador/{PedidoID}")]
        [HttpGet]
        public HttpResponseMessage ListarSeccaoUtilizador(string PedidoID)
        {
            try
            {
                string query = $@"SELECT ID,Seccao,Utilizador FROM STP_Processos Where ID = '{PedidoID}'";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma seção encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter seções: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarPedidos")]
        [HttpGet]
        public HttpResponseMessage ListarPedidos()
        {
            try
            {
                string query = @"SELECT 
                                sp.ID, 
                                sp.Cliente,
	                            c.Nome AS Nome, 
                                sp.DataHoraAbertura, 
                                sp.Prioridade, 
                                sp.Estado, 
                                sp.NumProcesso,
								sp.DescricaoProb,
								sp.serie,
								sp.Tecnico
                            FROM 
                                STP_Processos sp
                            JOIN 
                                Clientes c ON sp.cliente = c.Cliente
                            ORDER BY 
                                sp.NumProcesso DESC";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma seção encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter seções: {ex.Message}");
            }
        }

        [Authorize]
        [Route("DaUltimoEstadoPedido/{PedidoID}")]
        [HttpGet]
        public HttpResponseMessage DaUltimoEstadoPedido(string PedidoID)
        {
            try
            {
                string query = $@"SELECT TOP 1 Estado FROM STP_Intervencoes WHERE ProcessoID = '{PedidoID}' ORDER BY Interv desc;";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma seção encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter seções: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarIntervencoes/{PedidoId}")]
        [HttpGet]
        public HttpResponseMessage ListarIntervencoes(string PedidoId)
        {
            try
            {
                string query = $@"
                                SELECT 
                                    I.ID,
                                    I.Interv,
                                    I.DataHoraInicio, 
                                    I.DataHoraFim,
                                    I.TipoInterv,
                                    I.Duracao,
                                    T.Nome,
                                    
                                    COALESCE(I.DescricaoResp, 'sem informação') AS DescricaoResp
                                FROM 
                                    STP_Intervencoes I 
                                LEFT JOIN
                                    STP_Tecnicos T ON I.Tecnico = T.Tecnico
                                WHERE ProcessoID =  '{PedidoId}'
                                order by Interv asc
                                ";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma seção encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter seções: {ex.Message}");
            }
        }
        [Authorize]
        [Route("CriarPedido")]
        [HttpPost]
        public HttpResponseMessage CriarPedido([FromBody] CriarPedidoRequest pedidoRequest)
        {
            if (pedidoRequest == null)
            {
                return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados de pedido não podem ser nulos.");
            }

            try
            {
                StpBEProcesso pedido = new StpBEProcesso
                {
                    Cliente = pedidoRequest.Cliente,
                    DescricaoObjecto = pedidoRequest.DescricaoObjecto,
                    DescricaoProblema = pedidoRequest.DescricaoProblema,
                    Origem = pedidoRequest.Origem,
                    TipoProcesso = pedidoRequest.TipoProcesso,
                    Prioridade = pedidoRequest.Prioridade,
                    Tecnico = pedidoRequest.Tecnico,
                    ObjectoID = pedidoRequest.ObjectoID,
                    TipoDoc = pedidoRequest.TipoDoc,
                    ID = Guid.NewGuid().ToString(),
                    Filial = "000",
                    Serie = pedidoRequest.Serie,
                    Estado = pedidoRequest.Estado,
                    Seccao = pedidoRequest.Seccao,
                    ContratoID = pedidoRequest.ContratoID,
                    ComoReproduzir = pedidoRequest.ComoReproduzir,
                    Contacto = pedidoRequest.Contacto,
                    DataHoraAbertura = pedidoRequest.datahoraabertura,
                    DataHoraFimPrevista = DateTime.Now.AddDays(30)
                };

                StpBEPedido pedido1 = new StpBEPedido
                {
                    ID = pedido.ID,
                    Filial = "000",
                    TipoDoc = pedido.TipoDoc,
                    Serie = pedido.Serie,
                    Numero = pedido.NumProcesso,
                    Pedido = pedido.Processo,
                    RecebidoPor = pedido.Tecnico,
                    Origem = pedido.Origem,
                    Utilizador = null,
                    Cliente = pedido.Cliente,
                    Contacto = pedido.Contacto,
                    DataHora = pedido.DataHoraAbertura

                };

                string query = $@"INSERT INTO STP_LinhasPedido(NumLinha, ProcessoID, PedidoID) VALUES (1, '{pedido1.ID}' , '{pedido1.ID}');";



                ProductContext.MotorLE.ServicosTecnicos.Processos.Actualiza(pedido);
                ProductContext.MotorLE.ServicosTecnicos.Pedidos.Actualiza(pedido1);
                ProductContext.MotorLE.DSO.ExecuteSQL(query);

                return Request.CreateResponse(HttpStatusCode.OK, "Pedido criado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao criar pedido: {ex.Message}");
            }
        }












        //TODO

        [Authorize]
        [Route("MudaEstadoPedido/{EstadoID}")]
        [HttpGet]
        public HttpResponseMessage MudaEstadoPedido(string EstadoID, string ID)
        {
            try
            {
                string query = $@"UPDATE stp_processos SET estado = '{EstadoID}' WHERE id = '{ID}';";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }








        [Authorize]
        [Route("ListaProcessosTecnico/{TecnicoID}")]
        [HttpGet]
        public HttpResponseMessage ListaProcessosTecnico(string TecnicoID)
        {
            try
            {
                string query = $@"SELECT 
    i.*,
    p.*,
    c.Nome AS NomeCliente,
    co.PrimeiroNome AS NomeContacto,
    co.UltimoNome AS ApelidoContacto,
	con.TipoDoc
FROM 
    STP_Intervencoes i
LEFT JOIN 
    STP_Processos p ON i.ProcessoID = p.ID
LEFT JOIN 
    Clientes c ON p.cliente = c.Cliente
LEFT JOIN 
    Contactos co ON p.Contacto = co.Contacto
LEFT JOIN 
    PCM_Contratos con ON p.ContratoID = con.Id
WHERE 
    i.Tecnico = '{TecnicoID}';

                            ";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }


        [Authorize]
        [Route("ListaIntervencoesTecnico/{TecnicoID}")]
        [HttpGet]
        public HttpResponseMessage ListaIntervencoesTecnico(string TecnicoID)
        {
            try
            {
                string query = $@"SELECT * FROM STP_Intervencoes WHERE Tecnico = '{TecnicoID}'";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum objeto encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }






        [Authorize]
        [Route("CriarPedido/{cliente}/{descricaoObjecto}/{descricaoProblema}/{origem}/{tipoProcesso}/{prioridade}/{tecnico}/{objectoID}/{tipoDoc}/{serie}/{estado}/{seccao}/{comoReproduzir?}/{contacto?}")]
        [HttpPost]
        public HttpResponseMessage CriarPedido(string cliente, string descricaoObjecto, string descricaoProblema, string origem, string tipoProcesso, string prioridade, string tecnico, string objectoID, string tipoDoc, string serie, string estado, string seccao, string comoReproduzir = null, string contacto = null)
        {
            try
            {
                StpBEProcesso pedido = new StpBEProcesso
                {
                    Cliente = cliente,
                    DescricaoObjecto = descricaoObjecto,
                    DescricaoProblema = descricaoProblema,
                    Origem = origem,
                    TipoProcesso = tipoProcesso,
                    Prioridade = prioridade,
                    Tecnico = tecnico,
                    ObjectoID = objectoID,
                    TipoDoc = tipoDoc,
                    ID = Guid.NewGuid().ToString(),
                    Filial = "000",
                    Serie = serie,
                    Estado = estado,
                    Seccao = seccao,
                    ComoReproduzir = comoReproduzir,
                    Contacto = contacto,
                    DataHoraAbertura = DateTime.Now,
                    DataHoraFimPrevista = DateTime.Now.AddDays(30)
                };

                ProductContext.MotorLE.ServicosTecnicos.Processos.Actualiza(pedido);

                return Request.CreateResponse(HttpStatusCode.OK, "Pedido criado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao criar pedido: {ex.Message}");
            }
        }









        [Authorize]
        [Route("CriarIntervencoes")]
        [HttpPost]
        public HttpResponseMessage CriarIntervencoes([FromBody] CriarIntervencaoRequest request)
        {
            try
            {
                int ultimaIntervencao = 1;
                try
                {
                    ultimaIntervencao = ProductContext.MotorLE.ServicosTecnicos.Intervencoes.DaNumeroUltimaIntervencao(request.ProcessoID);
                    ultimaIntervencao++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao obter a última intervenção: {ex.Message}");
                }

                // Criar nova intervenção
                StpBEIntervencao intervencao = new StpBEIntervencao
                {
                    ID = Guid.NewGuid().ToString(),
                    ProcessoID = request.ProcessoID,
                    TipoIntervencao = request.TipoIntervencao,
                    Duracao = request.Duracao,
                    DuracaoReal = request.DuracaoReal,
                    DataHoraInicio = request.DataHoraInicio,
                    DataHoraFim = request.DataHoraFim,
                    Tecnico = request.Tecnico,
                    EstadoAnt = request.EstadoAnt,
                    Estado = request.Estado,
                    SeccaoAnt = request.SeccaoAnt,
                    Seccao = request.Seccao,
                    Moeda = "EUR",
                    Cambio = 1,
                    Utilizador = request.Utilizador,
                    NumIntervencao = ultimaIntervencao,
                    DescricaoResposta = string.IsNullOrEmpty(request.DescricaoResposta) ? null : request.DescricaoResposta,
                    ImputadoContrato = true,
                };

                // Obter custo e preço do técnico
                string queryTecnico = $@"SELECT * FROM STP_Tecnicos WHERE Tecnico = '{request.Tecnico}'";
                var bdTecnico = ProductContext.MotorLE.Consulta(queryTecnico);
                var horacusto = bdTecnico.DaValor<double>("HoraCusto");
                var horapreco = bdTecnico.DaValor<double>("HoraPreco");

                // Obter contrato
                string queryContrato = $@"
            SELECT DISTINCT 
                p.ContratoID,
					c.ID,
					p.Cliente,
					c.HorasTotais, 
					c.HorasGastas 
            FROM 
                STP_Processos p
            LEFT JOIN 
                STP_Contratos c
            ON 
                p.ContratoID = c.ID
            WHERE 
                p.ID = '{request.ProcessoID}'";

                var bdContrato = ProductContext.MotorLE.Consulta(queryContrato);
                var horasTotais = bdContrato.DaValor<object>("HorasTotais");
                var horasGastas = bdContrato.DaValor<object>("HorasGastas");

                double horasTotaisDouble = horasTotais != null ? Math.Round(Convert.ToDouble(horasTotais), 2) : 0;
                double horasGastasDouble = horasGastas != null ? Math.Round(Convert.ToDouble(horasGastas), 2) : 0;




                foreach (var artigo in request.Artigos)
                {
                    // Converter duração para horas
                    double duracaoEmHoras = Math.Round((double)intervencao.Duracao / 60, 2);

                    StpEstadoFacturacaoLinha artigoFaturacao;
                    if (horasTotaisDouble > horasGastasDouble)
                    {
                        artigoFaturacao = StpEstadoFacturacaoLinha.NaoFacturar;
                    }
                    else
                    {
                        artigoFaturacao = StpEstadoFacturacaoLinha.Facturar;
                    }





                    // Criar artigo intervenção
                    StpBEArtigoIntervencao artigoIntervencao = new StpBEArtigoIntervencao
                    {
                        ID = Guid.NewGuid().ToString(),
                        Artigo = artigo.artigo ?? "SERV.CONS",
                        Descricao = artigo.descricao ?? "ValorDefaultDescricao",
                        Unidade = artigo.unidade ?? "",
                        Armazem = artigo.armazem ?? "",
                        QtdeCusto = (double)duracaoEmHoras,
                        PrecoCusto = horacusto,
                        QtdeCliente = (double)duracaoEmHoras,
                        PrecoCliente = horapreco,
                        DescontoCliente = artigo.descontoCliente ?? 0,
                        EstadoFacturacao = artigoFaturacao,
                        TipoLinha = "20",
                    };

                    intervencao.ArtigosIntervencao.Insere(artigoIntervencao);

                }

                // Registar intervenção
                ProductContext.MotorLE.ServicosTecnicos.Intervencoes.Actualiza(intervencao);

                var updateTipoServico = $@"update STP_ArtigosIntervencao
                                            set tiposervico = 1
                                            from  STP_ArtigosIntervencao
                                            where IntervencaoID='{intervencao.ID}';";
                ProductContext.MotorLE.DSO.ExecuteSQL(updateTipoServico);


                var getValores = $@"	SELECT 
                                        i.ProcessoID,
                                        p.ContratoID,
                                        c.HorasGastas,
                                        i.ID, 
                                        p.*, 
                                        c.*
                                    FROM 
                                        STP_Intervencoes i
                                    JOIN 
                                        STP_Processos p ON i.ProcessoID = p.ID
                                    JOIN 
                                        STP_Contratos c ON p.ContratoID = c.ID
                                    WHERE 
                                        i.ID = '{intervencao.ID}'";

                var Data = ProductContext.MotorLE.Consulta(getValores);
                //var contratoID = Data.DaValor<object>("ContratoID");
                var contrato = Data.DaValor<string>("ContratoID");
                //  var cID = contratoID != null ? Convert.ToString(contratoID) : "";
                var hgD = horasGastas != null
                                        ? Math.Round(Convert.ToDecimal(horasGastas), 2)
                                        : 0;
                var cont = contrato != null ? contrato.ToString() : "";
                var du = Math.Round((decimal)intervencao.Duracao / 60, 2);

                var calculo = hgD + du;
                var resultado = calculo.ToString(CultureInfo.InvariantCulture);
                // Gera o comando SQL sem aspas ao redor do valor decimal
                var dataUpdate = $@"UPDATE STP_Contratos SET HorasGastas = {resultado} FROM STP_Contratos WHERE ID = '{cont}'";



                ProductContext.MotorLE.DSO.ExecuteSQL(dataUpdate);


                return Request.CreateResponse(HttpStatusCode.OK, $"{ultimaIntervencao}");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao criar intervenção: {ex.Message}");
            }
        }


        [Authorize]
        [Route("EditaId")]
        [HttpGet]
        public HttpResponseMessage EditaId()
        {
            try
            {
                var processo = new StpBEProcesso
                {
                    Cliente = "LIMA",
                    DescricaoObjecto = "ASSISTENCIA",
                    DescricaoProblema = "teste",
                    Origem = "TEL",
                    TipoProcesso = "HD",
                    Prioridade = "BX",
                    Tecnico = "JTA",
                    ObjectoID = "d1234ce3-8655-11ef-af1a-b047e94cac6a",
                    TipoDoc = "PA",
                    ID = Guid.NewGuid().ToString(),
                    Filial = "000",
                    Serie = "2024",
                    Estado = "0",
                    Seccao = "FU"
                };

                var intervencao = new StpBEIntervencao
                {
                    ID = Guid.NewGuid().ToString(),
                    ProcessoID = processo.ID,
                    TipoIntervencao = "REMOTO",
                    Duracao = 1,
                    DuracaoReal = 1,
                    DataHoraInicio = DateTime.Now,
                    DataHoraFim = DateTime.Now.AddDays(30),
                    Tecnico = "JTA",
                    EstadoAnt = "1",
                    Estado = "1",
                    SeccaoAnt = "DV",
                    Seccao = "DV",
                    Moeda = "EUR",
                    Cambio = 1,
                    Utilizador = "jtalm"
                };

                var linhaIntervencao = new StpBEArtigoIntervencao();

                var response = ProductContext.MotorLE.ServicosTecnicos.Intervencoes.SugereLinhaIntervencao(ref linhaIntervencao, intervencao, processo, "A0001", "teste", 20, "UN", "s", 0, 0, 0, "A4", "A4");

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao editar ID: {ex.Message}");
            }
        }

        [Authorize]
        [Route("FechaProcessoID/{Id}")]
        [HttpPost]

        public HttpResponseMessage FechaProcessoID(string Id)
        {
            try
            {
                string dataHoraAtual = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                string query = $@"
            UPDATE stp_processos 
            SET fechado = 1, Estado = 0, datahorafecho = '{dataHoraAtual}'
            WHERE id = '{Id}';";

                ProductContext.MotorLE.DSO.ExecuteSQL(query);

                return Request.CreateResponse(HttpStatusCode.OK, "Fechado");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter objetos: {ex.Message}");
            }
        }







    }

    [RoutePrefix("Obras")]
    public class ObrasController : ApiController
    {
        [Authorize]
        [Route("ListaObras")]
        [HttpGet]
        public HttpResponseMessage ListaObras()
        {
            try
            {
                string query = @"
            SELECT COP_Obras.ID, CASE COP_Obras.Tipo WHEN 'O' THEN 'Obra' WHEN 'C' THEN 'Contrato Adicional' WHEN 'S' THEN 'Subempreitada' ELSE '' END AS Tipo, COP_Obras.Codigo, COP_Obras.Descricao, CASE (COP_Obras.Estado)     WHEN 'ADJU' THEN 'Adjudicado'     WHEN 'CANC' THEN 'Cancelado/Recusado'     WHEN 'CONC' THEN 'Concurso/Orçamentado'     WHEN 'CONS' THEN 'Consignado/Em Curso'     WHEN 'PERD' THEN 'Perdido'     WHEN 'SUSP' THEN 'Suspenso'     WHEN 'RECD' THEN 'Receção Definitiva/Fechado'     WHEN 'RECP' THEN 'Receção Provisória/Terminado'     ELSE '' END  AS Estado, COP_Obras.Titulo, COP_Obras.[Local], Geral_Zona.Desig AS Zona, COP_Obras.Tipo as LTipo, COP_Obras.Estado as LEstado , COP_Obras.DataProposta, COP_Obras.NumConcurso, COP_Obras.DataDRepublica, COP_Obras.PrecoBase, COP_Obras.Consorcio, COP_Obras.PrazoExecucao, COP_Obras.PrazoFixado, COP_Obras.DataEntregaPropostas, COP_Obras.DataAberturaPropostas, COP_TipoEmpreitadas.Descricao AS TipoEmpreitada, COP_FormaContratos.Descricao AS FormaContratos, COP_TipoPropostas.Descricao AS TipoProposta,COP_SituacoesObra.Descricao as Situacao, Geral_Entidade.Codigo AS CodEntA, Geral_Entidade.Nome AS EntidadeA, CASE COP_Obras.TipoEntidadeA WHEN '1' THEN 'Dono da Obra' WHEN '2' THEN 'Empreiteiro' WHEN '3' THEN 'Subempreiteiro' WHEN '5' THEN 'Fiscalização' WHEN '6' THEN 'Coordenação de Segurança' WHEN '4' THEN 'Projetista' WHEN '7' THEN 'Fornecedor' ELSE '' END AS 
TipoEntA, CASE COP_Obras.ERPTipoEntidadeA WHEN 'C' THEN 'Cliente' WHEN 'F' THEN 'Fornecedor' WHEN 'R' THEN 'Credor' WHEN 'D' 
THEN 'Devedor' ELSE '' END as TipoEntidadeERP, COP_Obras.ERPEntidadeA AS EntidadeERP, V_Entidades.Nome AS NomeEntERP, 
V_Entidades.Morada, V_Entidades.Localidade, V_Entidades.Cp, V_Entidades.CpLoc, V_Entidades.Tel, V_Entidades.Fax, 
ISNULL(Geral_Entidade.NIPC,V_Entidades.NumContrib) AS NIPC, COP_Obras.SujRevisao, COP_Obras.CriadoPor, COP_Obras.Utilizador, 
COP_Obras.DataUltimaActualizacao, COP_Obras.Notas, CASE WHEN COP_Obras.ObraPaiID IS NULL THEN '' ELSE O2.Codigo + ' - '+ 
O2.Descricao END AS ObraPai, T2.Descricao TipoContratoAdicional, COP_Obras.CADefault CAPorDefeito ,COP_Obras.DataCriacao , 
COP_Obras.Freguesia, COP_Obras.Conselho, COP_Obras.ValorAdjudicacao, COP_TipoObras.Descricao AS TipoObra, 
COP_Empreitadas.Descricao AS Empreitada, COP_Obras.RespSeguranca, COP_Obras.Encarregado, COP_Obras.Apontador, 
COP_Obras.PorAmortizar, COP_Obras.DataAdjudicacao, COP_Obras.DataContrato, COP_Obras.DataTribunalContas,  
COP_Obras.Engenheiro, COP_Obras.DataConsignacao, COP_Obras.DataConclusaoPrevista, COP_Obras.DataConclusaoCorrigida, 
COP_Obras.DataRecepcaoProvisoria, COP_Obras.DataCaucoes, COP_Obras.PrazoExtincaoCaucoes, COP_Obras.DataRecepcaoDefinitiva, 
COP_Obras.PrazoGarantiaRecDefinitiva, COP_Obras.DataEntregaObra, COP_Obras.DataPrevisao1Entrega, 
COP_Obras.DataPrevisao2Entrega, COP_Obras.CalculoPrazos, COP_Obras.MultaDiaria, COP_Obras.TotalDiario, 
COP_Obras.PremioDiario, COP_Obras.TotalPremio, (SELECT isnull(SUM(Valor),0) Valor FROM COP_RParcelaRevisoes         WHERE 
RParcelaID IN (SELECT ID FROM COP_RParcelas WHERE AutoID IN (SELECT ID AS AutoID FROM COP_Autos WHERE ObraID = 
COP_Obras.ID))) as ValorRevisao,Geral_Entidade_1.Codigo AS CodEntB, Geral_Entidade_1.Nome AS EntidadeB, CASE 
COP_Obras.TipoEntidadeB WHEN '1' THEN 'Dono da Obra' WHEN '2' THEN 'Empreiteiro' WHEN '3' THEN 'Subempreiteiro' WHEN '5' THEN 'Fiscalização' WHEN '6' THEN 'Coordenação de Segurança' WHEN '4' THEN 'Projetista' WHEN '7' THEN 'Fornecedor' ELSE '' END AS TipoEntB, CASE COP_Obras.ERPTipoEntidadeB WHEN 'C' THEN 'Cliente' WHEN 'F' THEN 'Fornecedor' WHEN 'R' THEN 'Credor' WHEN 'D' THEN 'Devedor' ELSE '' END as TipoEntidadeERP_B, COP_Obras.ERPEntidadeB AS EntidadeERP_B, V_Entidades_B.Nome AS NomeEntERP_B, V_Entidades_B.Morada as Morada_B, V_Entidades_B.Localidade AS Localidade_B, V_Entidades_B.CP AS Cp_B, V_Entidades_B.CpLoc AS CpLoc_B, V_Entidades_B.Tel as Tel_B, V_Entidades_B.Fax as Fax_B, ISNULL(Geral_Entidade_1.NIPC,V_Entidades_B.NumContrib) AS NIPC_B FROM COP_Obras LEFT OUTER JOIN Geral_Zona ON COP_Obras.ZonaID = Geral_Zona.ZonaId LEFT OUTER JOIN COP_FormaContratos ON COP_Obras.TipoCont = COP_FormaContratos.Tipo LEFT OUTER JOIN COP_TipoEmpreitadas ON COP_Obras.TipoEmp = COP_TipoEmpreitadas.Tipo LEFT OUTER JOIN COP_TipoPropostas ON COP_Obras.TipoProp = COP_TipoPropostas.Tipo LEFT OUTER JOIN COP_SituacoesObra ON COP_Obras.SituacaoID = COP_SituacoesObra.ID LEFT OUTER JOIN Geral_Entidade ON COP_Obras.EntidadeIDA = Geral_Entidade.EntidadeId LEFT JOIN V_Entidades ON COP_Obras.ERPTipoEntidadeA=V_Entidades.TipoEntidade AND COP_Obras.ERPEntidadeA=V_Entidades.Entidade LEFT OUTER JOIN COP_Obras O2 ON COP_Obras.ObraPaiID = O2.ID LEFT OUTER JOIN COP_TiposContratoAdicional T2 ON COP_Obras.TipoContratoAdicionalID = T2.ID LEFT OUTER JOIN COP_TipoObras ON COP_Obras.TipoObra = COP_TipoObras.Tipo LEFT OUTER JOIN COP_Empreitadas ON COP_Empreitadas.ID = COP_Obras.EmpreitadaID LEFT OUTER JOIN Geral_Entidade Geral_Entidade_1 ON COP_Obras.EntidadeIDB = Geral_Entidade_1.EntidadeId LEFT JOIN V_Entidades V_Entidades_B ON COP_Obras.ERPTipoEntidadeB=V_Entidades_B.TipoEntidade AND COP_Obras.ERPEntidadeB=V_Entidades_B.Entidade WHERE COP_Obras.Projecto=0  AND (COP_Obras.Tipo ='O')  AND (COP_Obras.Estado IN ('ADJU', 'CONS')) 
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }



        [Authorize]
        [Route("GetCAdicionais_Estimado/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetCAdicionais_Estimado(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT 
                CAST(SUM(O.TotExec) + SUM(ISNULL(CI.Valor, 0)) AS DECIMAL(18,2)) AS Valor
            FROM 
                Orcamentos_Orcamento O
            INNER JOIN 
                COP_Obras CA ON CA.ID = O.ObraID
            LEFT JOIN (
                SELECT 
                    CA.ID, 
                    CAST(SUM(ISNULL(OI.QuantExec, 0) * ISNULL(OI.PrecoExecFormula, 0)) AS DECIMAL(18,2)) AS Valor
                FROM 
                    Orcamentos_Item OI
                INNER JOIN 
                    Orcamentos_Orcamento O ON O.OrcId = OI.OrcId
                INNER JOIN 
                    COP_Obras CA ON CA.ID = O.ObraID
                WHERE 
                    CA.Tipo = 'C' AND 
                    CA.ObraPaiID = '{IdObra}' AND 
                    CA.Estado NOT IN ('CONC','CANC','PERD') AND 
                    CA.DataAdjudicacao <= '{data}' AND 
                    O.OrcPrincipal = 1 AND 
                    OI.Classificacao IN (3, 6) AND 
                    OI.IsGroup = 0
                GROUP BY 
                    CA.ID
                UNION
                SELECT 
                    CA.ID, 
                    CAST(SUM(ISNULL(CI.QuantExec, 0) * ISNULL(CIV.Diverso, 0)) AS DECIMAL(18,2)) AS Valor
                FROM 
                    Orcamentos_CustIndValores CIV
                INNER JOIN 
                    Orcamentos_CustInd CI ON CI.CustIndID = CIV.CustIndID
                INNER JOIN 
                    Orcamentos_Orcamento O ON O.OrcId = CI.OrcId
                INNER JOIN 
                    COP_Obras CA ON CA.ID = O.ObraID
                WHERE 
                    CA.Tipo = 'C' AND 
                    CA.ObraPaiID = '{IdObra}' AND 
                    CA.Estado NOT IN ('CONC','CANC','PERD') AND 
                    CA.DataAdjudicacao <= '{data}' AND 
                    O.OrcPrincipal = 1 AND 
                    CIV.Modo = 2
                GROUP BY 
                    CA.ID
            ) AS CI ON CI.ID = CA.ID
            WHERE 
                CA.Tipo = 'C' AND 
                CA.ObraPaiID = '{IdObra}' AND 
                CA.Estado NOT IN ('CONC','CANC','PERD') AND 
                CA.DataAdjudicacao <= '{data}' AND 
                O.OrcPrincipal = 1 AND 
                O.TotExec <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetSubempreitadas_Real/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetSubempreitadas_Real(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT SUM(CS.Valor) AS Valor
            FROM FN_COP_DaObrasMaisContratosAdicionais('CANC, CONC, PERD') OCA
            CROSS APPLY FN_COP_DaCustosSubempreitadas(OCA.ID, 1, NULL, '{data}') CS
            WHERE OCA.Tipo IN ('O', 'C') 
              AND OCA.ObraPrincipalID = '{IdObra}'
            GROUP BY CS.ID, CS.Codigo, CS.Descricao, OCA.Codigo, OCA.Descricao, OCA.Tipo
            ORDER BY CS.Codigo DESC;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetOutrosCustos_Real/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetOutrosCustos_Real(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT CAST(SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) AS DECIMAL(11,2)) AS Valor
            FROM COP_MovimentosCusto C1
            INNER JOIN COP_Obras O ON C1.ObraID = O.ID
            WHERE (O.ID = '{IdObra}' 
                OR (O.ObraPaiID = '{IdObra}' AND O.Tipo = 'C'))
                AND C1.Data <= '{data}'
                AND C1.Origem IN ('C', 'V', 'M', 'D', 'S', 'B', 'F') 
                AND O.Tipo <> 'S'
            HAVING SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetFichasPessoal_Real/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetFichasPessoal_Real(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601
                string query = $@"
            SELECT CAST(SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) AS DECIMAL(11,2)) AS Valor
            FROM COP_MovimentosCusto C1
                INNER JOIN COP_Obras O ON C1.ObraID = O.ID
            WHERE (C1.ObraID = '{IdObra}' 
                OR (O.ObraPaiID = '{IdObra}' AND O.Tipo = 'C'))
                AND C1.Data <= '{data}'
                AND C1.Origem = 'P'
            HAVING SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetFichasEquipamento_Real/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetFichasEquipamento_Real(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601

                string query = $@"
            SELECT CAST(SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) AS DECIMAL(11,2)) AS Valor
            FROM COP_MovimentosCusto C1
                INNER JOIN COP_Obras O ON C1.ObraID = O.ID
            WHERE (C1.ObraID = '{IdObra}' 
                OR (O.ObraPaiID = '{IdObra}' AND O.Tipo = 'C'))
                AND C1.Data <= '{data}'
                AND C1.Origem = 'E'
            HAVING SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetCustosManuais_Real/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetCustosManuais_Real(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT CAST(SUM(ISNULL(C.Quantidade, 0) * ISNULL(C.PrecoUnit, 0)) AS DECIMAL(11,2)) AS Valor
            FROM COP_Custos C
            INNER JOIN COP_Obras O ON C.ObraID = O.ID
            LEFT JOIN COP_Documentos D ON C.DocumentoID = D.Id
            WHERE (O.ID = '{IdObra}' 
                OR (O.ObraPaiID = '{IdObra}' AND O.Tipo = 'C')) 
                AND C.Data <= '{data}' 
                AND O.Tipo <> 'S'
            HAVING SUM(ISNULL(C.Quantidade, 0) * ISNULL(C.PrecoUnit, 0)) <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetTrabalhosMenos_Real/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetTrabalhosMenos_Real(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT CAST(SUM(ISNULL(af1.TMenosValorExec, 0)) AS DECIMAL(18,2)) AS Valor
            FROM COP_Autos af1
            INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
            WHERE (af1.ObraID = '{IdObra}' 
                OR (of1.ObraPaiID = '{IdObra}' AND of1.Tipo = 'C'))
                AND af1.Data <= '{data}' 
                AND ISNULL(af1.TMenosValorExec, 0) <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetSubempreitadas_Pendentes/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetSubempreitadas_Pendentes(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT CAST(SUM(CS.Valor) AS DECIMAL(18,2)) AS Valor
            FROM FN_COP_DaObrasMaisContratosAdicionais('CANC, CONC, PERD') OCA
            CROSS APPLY FN_COP_DaCustosSubempreitadas(OCA.ID, 0, NULL, '{data}') CS
            WHERE OCA.Tipo IN ('O', 'C') 
            AND OCA.ObraPrincipalID = '{IdObra}';
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetCAdicionais_Proveitos/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetCAdicionais_Proveitos(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT 
                CAST(SUM(orc1.TotVenda) AS DECIMAL(18,2)) AS Valor
            FROM Orcamentos_Orcamento orc1
            INNER JOIN COP_Obras of1 ON orc1.ObraID = of1.ID
            WHERE of1.Tipo = 'C' 
            AND of1.ObraPaiID = '{IdObra}' 
            AND of1.Estado NOT IN ('CONC','CANC','PERD')
            AND of1.DataAdjudicacao <= '{data}' 
            AND orc1.OrcPrincipal = 1 
            AND orc1.TotVenda <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetNaoAutorizados_Faturacao/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetNaoAutorizados_Faturacao(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT 
                CAST(SUM(
                    (CASE WHEN af1.Estado = 0 THEN af1.amvalor ELSE 0 END) + 
                    (CASE WHEN af1.EstadoTM = 0 THEN af1.tmvalor ELSE 0 END)
                ) AS DECIMAL(18,2)) AS Valor
            FROM COP_Autos af1
            INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
            WHERE (
                (of1.Tipo = 'C' AND of1.ObraPaiID = '{IdObra}' AND of1.Estado NOT IN ('CONC','CANC','PERD')) 
                OR (of1.Tipo <> 'S' AND of1.ID = '{IdObra}')
            )
            AND af1.Data <= '{data}' 
            AND (af1.Estado = 0 OR af1.EstadoTM = 0) 
            AND (
                (CASE WHEN af1.Estado = 0 THEN af1.amvalor ELSE 0 END) + 
                (CASE WHEN af1.EstadoTM = 0 THEN af1.tmvalor ELSE 0 END)
            ) <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetAutorizados_Faturacao/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetAutorizados_Faturacao(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT 
                CAST(SUM(
                    (CASE WHEN af1.Estado = 1 AND af1.AMCabecDocID IS NULL THEN af1.amvalor ELSE 0 END) +
                    (CASE WHEN af1.EstadoTM = 1 AND af1.TMCabecDocID IS NULL THEN af1.tmvalor ELSE 0 END)
                ) AS DECIMAL(18,2)) AS Valor
            FROM COP_Autos af1
            INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
            WHERE (
                    (of1.Tipo = 'C' AND of1.ObraPaiID = '{IdObra}' AND of1.Estado NOT IN ('CONC','CANC','PERD')) 
                    OR (of1.Tipo <> 'S' AND of1.ID = '{IdObra}')
                )
                AND af1.Data <= '{data}' 
                AND (
                    (af1.Estado = 1 AND af1.AMCabecDocID IS NULL) 
                    OR (af1.EstadoTM = 1 AND af1.TMCabecDocID IS NULL)
                )
                AND (
                    (CASE WHEN af1.Estado = 1 AND af1.AMCabecDocID IS NULL THEN af1.amvalor ELSE 0 END) +
                    (CASE WHEN af1.EstadoTM = 1 AND af1.TMCabecDocID IS NULL THEN af1.tmvalor ELSE 0 END)
                ) <> 0;
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetFaturados_Faturacao/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetFaturados_Faturacao(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT 
                CAST(SUM(
                    (CASE WHEN A.Estado = 1 AND A.AMCabecDocID IS NOT NULL THEN A.AMValor ELSE 0 END) +
                    (CASE WHEN A.EstadoTM = 1 AND A.TMCabecDocID IS NOT NULL THEN A.TMValor ELSE 0 END)
                ) AS DECIMAL(18,2)) AS TotalValor
            FROM COP_Autos A WITH (NOLOCK)
            INNER JOIN COP_Obras O WITH (NOLOCK) ON O.ID = A.ObraID
            WHERE ((O.ID = '{IdObra}' AND O.Tipo <> 'S') 
                   OR (O.Tipo = 'C' AND O.ObraPaiID = '{IdObra}')) 
                AND A.Data <= '{data}'
                AND ((A.Estado = 1 AND A.AMCabecDocID IS NOT NULL) 
                     OR (A.EstadoTM = 1 AND A.TMCabecDocID IS NOT NULL))
                AND ((CASE WHEN A.Estado = 1 AND A.AMCabecDocID IS NOT NULL THEN A.AMValor ELSE 0 END) +
                    (CASE WHEN A.EstadoTM = 1 AND A.TMCabecDocID IS NOT NULL THEN A.TMValor ELSE 0 END) <> 0);
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetTrabalhosMenos_Faturacao/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetTrabalhosMenos_Faturacao(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT 
                CAST(SUM(ISNULL(af1.TMenosValor, 0)) AS DECIMAL(18,2)) AS TotalValor
            FROM COP_Autos af1
            INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
            WHERE (af1.ObraID = '{IdObra}' 
                   OR (of1.ObraPaiID = '{IdObra}' AND of1.Tipo = 'C')) 
              AND af1.Data <= '{data}' 
              AND ISNULL(af1.TMenosValor, 0) <> 0;
        ";
                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetFaturada_RevisaoPrecos/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetFaturada_RevisaoPrecos(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601;
                string query = $@"
            SELECT 
                CAST(SUM(ISNULL(F.ValorRevisao, 0)) AS DECIMAL(18,2)) AS TotalValor
            FROM COP_RevisaoFacturas F WITH (NOLOCK)
            INNER JOIN CabecDoc C WITH (NOLOCK) ON C.Id = F.CabecDocID
            WHERE F.ObraID = '{IdObra}' 
              AND F.Data <= '{data}' 
              AND (ISNULL(F.ValorRevisao, 0) <> 0 OR ISNULL(F.ValorFacturado, 0) <> 0);
        ";
                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetAutosMedicao_Execucao/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetAutosMedicao_Execucao(string IdObra)
        {
            try
            {
                var data = DateTime.Now.ToString("yyyy-MM-dd"); // Formata a data no padrão ISO 8601; // A data atual é obtida, mas não é utilizada na consulta.
                string query = $@"
            SELECT 
                Data,
                AMCodigo,
                AMValor,
                TMCodigo,
                TMValor,
                TMenosCodigo,
                TMenosValor,
                Estado 
            FROM COP_Autos 
            WHERE ObraID = '{IdObra}';
        ";
                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter obras: {ex.Message}");
            }
        }



        [Authorize]
        [Route("GetControlo/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetControlo(string IdObra)
        {
            try
            {
                string query = $@"
        SELECT T1.Ordem, T1.Tipo, ISNULL(T2.Valor, T1.Valor) Valor
FROM (SELECT 1 AS Ordem, 'P1' AS Tipo, 0 AS Valor) AS T1
    LEFT JOIN (
        SELECT 1 AS Ordem, 'P1' AS Tipo, ISNULL(TotVenda, 0) AS Valor 
        FROM Orcamentos_Orcamento
        WHERE ObraID = '{IdObra}' AND OrcPrincipal = 1
    ) AS T2 ON T2.Ordem = T1.Ordem
UNION
SELECT 2 AS Ordem, 'P2' AS Tipo, SUM(TotVenda) AS Valor
FROM Orcamentos_Orcamento orc1
    INNER JOIN COP_Obras of1 ON orc1.ObraId = of1.ID
WHERE of1.Tipo = 'C' AND of1.ObraPaiID = '{IdObra}' AND of1.Estado NOT IN ('CANC','CONC','PERD')
    AND of1.DataAdjudicacao <= GETDATE() AND orc1.OrcPrincipal = 1
UNION
SELECT 3 AS Ordem, 'P3' AS Tipo,
    SUM((CASE WHEN af1.Estado = 0 THEN af1.amvalor ELSE 0 END)
    + (CASE WHEN af1.EstadoTM = 0 THEN af1.tmvalor ELSE 0 END)) AS Valor
FROM COP_Autos af1
    INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
WHERE ((of1.Tipo = 'C' AND of1.ObraPaiID = '{IdObra}' AND of1.Estado NOT IN ('CANC','CONC','PERD')) 
    OR (of1.Tipo <> 'S' AND of1.ID = '{IdObra}'))
    AND af1.Data <= GETDATE() AND (af1.Estado = 0 OR af1.EstadoTM = 0)
UNION
SELECT 4 AS Ordem, 'P4' AS Tipo,
    SUM((CASE WHEN af1.Estado = 1 AND af1.AMCabecDocID IS NULL THEN af1.amvalor ELSE 0 END)
    + (CASE WHEN af1.EstadoTM = 1 AND af1.TMCabecDocID IS NULL THEN af1.tmvalor ELSE 0 END)) AS Valor
FROM COP_Autos af1
    INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
WHERE (((of1.Tipo = 'C' AND of1.ObraPaiID = '{IdObra}') AND (of1.Estado NOT IN ('CANC','CONC','PERD'))) 
    OR (of1.Tipo <> 'S' AND of1.ID = '{IdObra}'))
    AND af1.Data <= GETDATE() AND ((af1.Estado = 1 AND af1.AMCabecDocID IS NULL) OR (af1.EstadoTM = 1 AND af1.TMCabecDocID IS NULL))
UNION
SELECT 5 AS Ordem, 'P5' AS Tipo,
    SUM((CASE WHEN af1.Estado = 1 AND af1.AMCabecDocID IS NOT NULL THEN af1.amvalor ELSE 0 END)
    + (CASE WHEN af1.EstadoTM = 1 AND af1.TMCabecDocID IS NOT NULL THEN af1.tmvalor ELSE 0 END)) AS Valor
FROM COP_Autos af1
    INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID 
WHERE ((of1.Tipo = 'C' AND of1.ObraPaiID = '{IdObra}') OR (of1.Tipo <> 'S' AND of1.ID = '{IdObra}'))
    AND af1.Data <= GETDATE() AND ((af1.Estado = 1 AND af1.AMCabecDocID IS NOT NULL) OR (af1.EstadoTM = 1 AND af1.TMCabecDocID IS NOT NULL))
UNION
SELECT 6 AS Ordem, 'P6' AS TIPO, ISNULL(SUM(af1.TMenosValor), 0) AS Valor
FROM COP_Autos af1
    INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
WHERE (ObraID = '{IdObra}' OR (of1.ObraPaiID = '{IdObra}' AND of1.Tipo = 'C')) AND af1.Data <= GETDATE()
UNION
SELECT 7 AS Ordem, 'P7' AS Tipo, ISNULL(SUM(RPR.Valor), 0) AS Valor
FROM COP_RParcelaRevisoes RPR
    INNER JOIN COP_RParcelas RP ON RPR.RParcelaID = RP.ID 
    INNER JOIN COP_Autos A ON A.ID = RP.AutoID 
WHERE A.ObraID = '{IdObra}' AND A.Data <= GETDATE()
    AND (YEAR(RPR.Mes) < YEAR(GETDATE()) OR (YEAR(RPR.Mes) = YEAR(GETDATE()) AND MONTH(RPR.Mes) <= MONTH(GETDATE())))
UNION
SELECT 8 AS Ordem, 'P8' AS Tipo, SUM(ISNULL(ValorFacturado, 0)) AS Valor
FROM COP_RevisaoFacturas RF
WHERE RF.ObraID = '{IdObra}' AND RF.Data <= GETDATE()
UNION
SELECT T1.Ordem, T1.Tipo, ISNULL(T2.Valor, T1.Valor) Valor
FROM (SELECT 9 AS Ordem, 'C1' AS Tipo, 0 AS Valor) AS T1
    LEFT JOIN (
        SELECT 9 AS Ordem, 'C1' AS Tipo, (CASE WHEN ISNULL(TotExec, 0) = 0 THEN ISNULL(TotCusto, 0) ELSE ISNULL(TotExec, 0) END) AS Valor
        FROM Orcamentos_Orcamento
        WHERE ObraID = '{IdObra}' AND OrcPrincipal = 1
    ) AS T2 ON T2.Ordem = T1.Ordem
UNION
SELECT 10 AS Ordem, 'C2' AS Tipo, SUM(TotExec) AS Valor
FROM Orcamentos_Orcamento orc1
    INNER JOIN COP_Obras of1 ON orc1.ObraId = of1.ID 
WHERE of1.Tipo = 'C' AND of1.ObraPaiID = '{IdObra}' AND of1.Estado NOT IN ('CANC','CONC','PERD')
    AND of1.DataAdjudicacao <= GETDATE() AND orc1.OrcPrincipal = 1
UNION
SELECT 11 AS Ordem, 'C3' AS Tipo, SUM(CS.Valor) AS Valor
FROM FN_COP_DaObrasMaisContratosAdicionais('CANC, CONC, PERD') OCA
    CROSS APPLY FN_COP_DaCustosSubempreitadas(OCA.ID, 1, NULL, GETDATE()) CS
WHERE OCA.Tipo IN ('O', 'C') AND OCA.ObraPrincipalID = '{IdObra}'
UNION
SELECT 12 AS Ordem, 'C4' AS Tipo, SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) AS Valor
FROM COP_MovimentosCusto C1
    INNER JOIN COP_Obras O ON C1.ObraID = O.ID
WHERE (O.ID = '{IdObra}' OR O.ObraPaiID = '{IdObra}') AND C1.Data <= GETDATE()
    AND C1.Origem IN ('C', 'V', 'M', 'D', 'S', 'B', 'F') AND O.Tipo <> 'S'
UNION
SELECT 13 AS Ordem, 'C5' AS Tipo, SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)) AS Valor
FROM COP_MovimentosCusto C1
    INNER JOIN COP_Obras O ON O.ID = C1.ObraID
WHERE (C1.ObraID = '{IdObra}' OR (O.ObraPaiID = '{IdObra}' AND O.Tipo = 'C'))
    AND C1.Data <= GETDATE() AND C1.Origem = 'P' AND O.Tipo <> 'S'
UNION
SELECT 14 AS Ordem, 'C6' AS Tipo, ISNULL(SUM(ISNULL(C1.Quantidade, 0) * ISNULL(C1.PrecoUnit, 0)), 0) AS Valor
FROM COP_MovimentosCusto C1
    INNER JOIN COP_Obras O ON O.ID = C1.ObraID
WHERE (C1.ObraID = '{IdObra}' OR (O.ObraPaiID = '{IdObra}' AND O.Tipo = 'C'))
    AND C1.Data <= GETDATE() AND C1.Origem = 'E' AND O.Tipo <> 'S'
UNION
SELECT 15 AS Ordem, 'C7' AS Tipo, SUM(ISNULL(C.Quantidade, 0) * ISNULL(C.PrecoUnit, 0)) AS Valor
FROM COP_Custos C
    INNER JOIN COP_Obras O ON C.ObraID = O.ID
    LEFT OUTER JOIN Geral_Classe ON C.ClasseID = Geral_Classe.ClasseId
    LEFT OUTER JOIN Orcamentos_Item OI ON C.ItemId = OI.ItemId
WHERE (O.ID = '{IdObra}' OR O.ObraPaiID = '{IdObra}') AND C.Data <= GETDATE() AND O.Tipo <> 'S'
UNION
SELECT 16 AS Ordem, 'C8' AS TIPO, ISNULL(SUM(af1.TMenosValorExec), 0) AS Valor
FROM COP_Autos af1
    INNER JOIN COP_Obras of1 ON af1.ObraID = of1.ID
WHERE (ObraID = '{IdObra}' OR (of1.ObraPaiID = '{IdObra}' AND of1.Tipo = 'C')) AND af1.Data <= GETDATE()
UNION
SELECT 17 AS Ordem, 'C9' AS Tipo, SUM(CS.Valor) AS Valor
FROM FN_COP_DaObrasMaisContratosAdicionais('CANC, CONC, PERD') OCA
    CROSS APPLY FN_COP_DaCustosSubempreitadas(OCA.ID, 0, NULL, GETDATE()) CS
WHERE OCA.Tipo IN ('O', 'C') AND OCA.ObraPrincipalID = '{IdObra}'
UNION
SELECT 18 AS Ordem, 'C10' AS Tipo, SUM(Valor) AS Valor
FROM (
    SELECT CAST(SUM(ISNULL(OI.QuantExec, 0) * ISNULL(OI.PrecoExecFormula, 0)) AS DECIMAL(18,2)) AS Valor
    FROM Orcamentos_Item OI
        INNER JOIN Orcamentos_Orcamento O ON O.OrcId = OI.OrcId
    WHERE O.ObraID = '{IdObra}' AND O.OrcPrincipal = 1 AND OI.Classificacao IN (3, 6) AND OI.IsGroup = 0
    UNION
    SELECT CAST(SUM(ISNULL(CI.QuantExec, 0) * ISNULL(CIV.Diverso, 0)) AS DECIMAL(18,2)) AS Valor
    FROM Orcamentos_CustIndValores CIV
        INNER JOIN Orcamentos_CustInd CI ON CI.CustIndID = CIV.CustIndID
        INNER JOIN Orcamentos_Orcamento O ON O.OrcId = CI.OrcId
    WHERE O.ObraID = '{IdObra}' AND O.OrcPrincipal = 1 AND CIV.Modo = 2
) AS CI
UNION
SELECT 19 AS Ordem, 'C11' AS Tipo, SUM(Valor) AS Valor
FROM (
    SELECT CAST(SUM(ISNULL(OI.QuantExec, 0) * ISNULL(OI.PrecoExecFormula, 0)) AS DECIMAL(18,2)) AS Valor
    FROM Orcamentos_Item OI
        INNER JOIN Orcamentos_Orcamento O ON O.OrcId = OI.OrcId
        INNER JOIN COP_Obras CA ON CA.ID = O.ObraID
    WHERE CA.Tipo = 'C' AND CA.ObraPaiID = '{IdObra}' AND CA.Estado NOT IN ('CANC','CONC','PERD')
        AND CA.DataAdjudicacao <= GETDATE() AND O.OrcPrincipal = 1 AND OI.Classificacao IN (3, 6) AND OI.IsGroup = 0
    UNION
    SELECT CAST(SUM(ISNULL(CI.QuantExec, 0) * ISNULL(CIV.Diverso, 0)) AS DECIMAL(18,2)) AS Valor
    FROM Orcamentos_CustIndValores CIV
        INNER JOIN Orcamentos_CustInd CI ON CI.CustIndID = CIV.CustIndID
        INNER JOIN Orcamentos_Orcamento O ON O.OrcId = CI.OrcId
        INNER JOIN COP_Obras CA ON CA.ID = O.ObraID
    WHERE CA.Tipo = 'C' AND CA.ObraPaiID = '{IdObra}' AND CA.Estado NOT IN ('CANC','CONC','PERD')
        AND CA.DataAdjudicacao <= GETDATE() AND O.OrcPrincipal = 1 AND CIV.Modo = 2
) AS CI
ORDER BY Ordem;

        ";
                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter controlo: {ex.Message}");
            }
        }


        [Authorize]
        [Route("GetFichasPessoal/{IdObra}")]
        [HttpGet]
        public HttpResponseMessage GetFichasPessoal(string IdObra)
        {
            try
            {
                string query = $@"
                    SELECT 
                        fp.ID,
                        fp.ObraID, 
                        CASE O.Tipo 
                            WHEN 'O' THEN 'Obra' 
                            WHEN 'S' THEN 'Subempreitada' 
                            WHEN 'C' THEN 'Contrato Adicional' 
                        END AS Tipo,
                        O.Codigo + '-' + O.Descricao AS Obra,
                        fp.CabecMovCBLID,
                        fp.Numero,
                        fp.Data,
                        fp.Notas,
                        fp.LigaCBL,
                        fp.CriadoPor,
                        fp.Validado,  
                        Movimentos.Diario,
                        Movimentos.NumDiario  
                    FROM 
                        COP_FichasPessoal fp  
                        INNER JOIN COP_Obras O ON fp.ObraID = O.ID 
                        LEFT OUTER JOIN Movimentos ON fp.CabecMovCBLID = Movimentos.ID 
                    WHERE 
                        O.ID = '{IdObra}' 
                        OR O.ObraPaiID = '{IdObra}';

        ";
                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter controlo: {ex.Message}");
            }
        }


        [Authorize]
        [Route("InsertPartesDiarias")]
        [HttpPost]
        public HttpResponseMessage InsertPartesDiarias([FromBody] ParteDiariaDto dados)
        {
            try
            {
                // Validação básica (opcional)
                if (dados == null || string.IsNullOrEmpty(dados.Utilizador) || dados.Numero <= 0)
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados inválidos.");
                }

                // Construir a consulta para COP_FichasPessoalItems
                string insertFichasPessoalItems = $@"
                INSERT INTO COP_FichasPessoalItems
                    ([ID], [FichasPessoalID], [ComponenteID], [Funcionario], [ClasseID], [Fornecedor], [SubEmpID], 
                     [NumHoras], [PrecoUnit], [TipoEntidade], [ColaboradorID], [SEPessoalID], [ManhaInicio], 
                     [ManhaFim], [TardeInicio], [TardeFim], [Data], [TotalHoras], [Integrado], [TipoHoraID], 
                     [FuncaoID], [ItemId]) 
                VALUES 
                    ('{Guid.NewGuid()}', '{dados.FichasPessoalID}', {dados.ComponenteID}, 
                     {(string.IsNullOrEmpty(dados.Funcionario) ? "NULL" : $"'{dados.Funcionario}'")}, {dados.ClasseID}, 
                     {(string.IsNullOrEmpty(dados.Fornecedor) ? "NULL" : $"'{dados.Fornecedor}'")}, 
                     {dados.SubEmpID}, {dados.NumHoras}, {dados.PrecoUnit}, '{dados.TipoEntidade}', 
                     {(dados.ColaboradorID.HasValue ? dados.ColaboradorID.ToString() : "NULL")}, 
                     {(string.IsNullOrEmpty(dados.SEPessoalID) ? "NULL" : $"'{dados.SEPessoalID}'")}, 
                     {(dados.ManhaInicio.HasValue ? $"'{dados.ManhaInicio.Value}'" : "NULL")}, 
                     {(dados.ManhaFim.HasValue ? $"'{dados.ManhaFim.Value}'" : "NULL")}, 
                     {(dados.TardeInicio.HasValue ? $"'{dados.TardeInicio.Value}'" : "NULL")}, 
                     {(dados.TardeFim.HasValue ? $"'{dados.TardeFim.Value}'" : "NULL")}, 
                     '{dados.Data:yyyy-MM-dd}', {dados.TotalHoras}, {dados.Integrado}, 
                     {(dados.TipoHoraID.HasValue ? dados.TipoHoraID.ToString() : "NULL")}, 
                     {(dados.FuncaoID.HasValue ? dados.FuncaoID.ToString() : "NULL")}, 
                     {(string.IsNullOrEmpty(dados.ItemId) ? "NULL" : $"'{dados.ItemId}'")})
                ";


                // Construir a consulta para COP_FichasPessoal
                string insertFichasPessoal = $@"
                INSERT INTO COP_FichasPessoal
                    ([ID], [Numero], [ObraID], [Data], [Encarregado], [Notas], [CabecMovCBLID], [LigaCBL],
                     [CriadoPor], [Utilizador], [DataUltimaActualizacao], [DocumentoID], [TipoEntidade], 
                     [SubEmpreiteiroID], [ColaboradorID], [Validado]) 
                VALUES 
                    ('{dados.ID}', {dados.Numero}, '{dados.ObraID}', '{dados.Data:yyyy-MM-dd}', 
                     {(string.IsNullOrEmpty(dados.Encarregado) ? "NULL" : $"'{dados.Encarregado}'")}, 
                     {(string.IsNullOrEmpty(dados.Notas) ? "NULL" : $"'{dados.Notas}'")}, 
                     {(string.IsNullOrEmpty(dados.CabecMovCBLID) ? "NULL" : $"'{dados.CabecMovCBLID}'")}, 
                     {dados.LigaCBL}, '{dados.CriadoPor}', '{dados.Utilizador}', '{dados.DataUltimaActualizacao:yyyy-MM-dd HH:mm:ss}', 
                     '{dados.DocumentoID}', '{dados.TipoEntidade}', 
                     {(string.IsNullOrEmpty(dados.SubEmpreiteiroID) ? "NULL" : $"'{dados.SubEmpreiteiroID}'")}, 
                     {(dados.ColaboradorID.HasValue ? dados.ColaboradorID.ToString() : "NULL")}, {dados.Validado})
                ";


                // Executar as consultas SQL
                ProductContext.MotorLE.DSO.ExecuteSQL(insertFichasPessoalItems);
                ProductContext.MotorLE.DSO.ExecuteSQL(insertFichasPessoal);

                return Request.CreateResponse(HttpStatusCode.OK, "Registos inseridos com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao inserir dados: {ex.Message}");
            }
        }

    }

    [RoutePrefix("Base")]
    public class BaseController : ApiController
    {
        [Authorize]
        [Route("LstClientes")]
        [HttpGet]
        public HttpResponseMessage LstClientes()
        {
            try
            {
                var response = ProductContext.MotorLE.Base.Clientes.LstClientes();

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum cliente encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter clientes: {ex.Message}");
            }
        }

    }

    [RoutePrefix("DataBases")]
    public class DataBasesController : ApiController
    {
        [Authorize]
        [Route("LstBaseDados")]
        [HttpGet]
        public HttpResponseMessage LstBaseDados()
        {
            try
            {
                string query = "SELECT name FROM sys.databases WHERE name NOT IN ('master','model','msdb')";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma baase dados encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter Base Dados: {ex.Message}");
            }
        }
    }

    [RoutePrefix("Artigo")]
    public class ArtigoController : ApiController
    {
        [Authorize]
        [Route("LstArtigos")]
        [HttpGet]
        public HttpResponseMessage LstArtigos()
        {
            try
            {
                var query = "SELECT * FROM Artigo";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum cliente encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter clientes: {ex.Message}");
            }
        }
    }

    [RoutePrefix("Imprimir")]
    public class ImprimirController : ApiController
    {
        [Authorize]
        [Route("Imprimir")]
        [HttpGet]
        public HttpResponseMessage Imprimir()
        {
            try
            {
                string pastaEspecifica = @"E:\Program Files"; // Altere para o caminho desejado
                string nomeArquivo = $"CS__{DateTime.Now:yyyyMMdd}.pdf";
                string caminhoCompleto = Path.Combine(pastaEspecifica, nomeArquivo);

                // Verificar se o diretório existe, caso contrário, cria o diretório
                if (!Directory.Exists(pastaEspecifica))
                {
                    Directory.CreateDirectory(pastaEspecifica);
                }

                // Inicializar e configurar o mapa
                ProductContext.Plataforma.Mapas.Inicializar("COP");

                ProductContext.Plataforma.Mapas.Destino = StdBSTipos.CRPEExportDestino.edFicheiro;
                ProductContext.Plataforma.Mapas.TipoFolha = StdBSTipos.CRPETipoFolha.tfA4;
                ProductContext.Plataforma.Mapas.SetFileProp(StdBSTipos.CRPEExportFormat.efPdf, caminhoCompleto);

                // Obter as informações do mapa
                var mapa = ProductContext.Plataforma.Mapas.GetMapaInfo("COP", "ORCADJ1");

                ProductContext.Plataforma.Mapas.MostraErros = true;

                ProductContext.Plataforma.Mapas.SetParametro("@OrcId", 3);
                ProductContext.Plataforma.Mapas.SetParametro("@Modo", 0);
                ProductContext.Plataforma.Mapas.SetParametro("@Depth", 99999);
                ProductContext.Plataforma.Mapas.SetParametro("@WithTotals", false);
                ProductContext.Plataforma.Mapas.SetParametro("@FilterGuid", string.Empty);
                // Gerar o arquivo PDF
                int imprimeMapa = ProductContext.Plataforma.Mapas.ImprimeListagem("ORCADJ1", "Documento", "P", 1, "N", "", 0, false, false);



                // Ler o arquivo gerado em bytes
                byte[] fileBytes = File.ReadAllBytes(caminhoCompleto);

                // Converter para base64 (se necessário, mas não é obrigatório para envio do arquivo)
                string fileString = Convert.ToBase64String(fileBytes);

                // Retornar o arquivo gerado como resposta
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);
                response.Content = new ByteArrayContent(fileBytes);
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
                response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                {
                    FileName = nomeArquivo
                };
                return response;



            }
            catch (Exception ex)
            {
                // Capturar e retornar o erro
                return Request.CreateResponse(HttpStatusCode.InternalServerError, new { message = $"Erro ao obter: {ex.Message}" });
            }
        }


    }

    [RoutePrefix("Word")]
    public class WordController : ApiController
    {
        [Authorize]
        [Route("Listar")]
        [HttpGet]
        public HttpResponseMessage Listar()
        {
            try
            {
                var query = "SELECt * FROM TDU_Oficios";
                var response = ProductContext.MotorLE.Consulta(query);


                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum registro encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter estados: {ex.Message}");
            }
        }

        [Authorize]
        [Route("Criar")]
        [HttpPost]
        public HttpResponseMessage Criar([FromBody] OficioDto novoOficio)
        {
            try
            {

                if (novoOficio == null)
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados inválidos.");
                }

                var Codigo = novoOficio.Codigo;
                var Assunto = novoOficio.Assunto;
                var Data = novoOficio.Data;
                var Remetente = novoOficio.Remetente;
                var Email = novoOficio.Email;
                var Texto1 = novoOficio.Texto1;
                var Texto2 = novoOficio.Texto2;
                var Template = novoOficio.Template;
                var Createdby = novoOficio.Createdby;
                var Texto3 = novoOficio.Texto3;
                var Obra = novoOficio.Obra;
                var DonoObra = novoOficio.donoObra;
                var Morada = novoOficio.Morada;
                var Localidade = novoOficio.Localidade;
                var CodPostal = novoOficio.CodPostal;
                var CodPostalLocal = novoOficio.CodPostalLocal;
                var Anexos = novoOficio.Anexos;
                var Texto4 = novoOficio.Texto4;
                var Texto5 = novoOficio.Texto5;
                var estado = novoOficio.Estado;
                var atencao = novoOficio.Atencao;

                string escapedTexto1 = Texto1?.Replace("'", "''");
                string escapedTexto2 = Texto2?.Replace("'", "''");
                string escapedTexto3 = Texto3?.Replace("'", "''");
                string query = $@"INSERT INTO TDU_Oficios (CDU_codigo, CDU_assunto, CDU_data, CDU_remetente, CDU_email, CDU_texto1, CDU_texto2, CDU_template, CDU_createdby, CDU_texto3, CDU_obra, CDU_DonoObra, CDU_Morada, CDU_Localidade, CDU_CodPostal, CDU_CodPostalLocal, CDU_Anexos, CDU_texto4, CDU_texto5, CDU_estado, CDU_AC) 
                 VALUES ('{Codigo}', '{Assunto}', '{Data}', '{Remetente}', '{Email}', '{Texto1}', '{Texto2}', '{Template}', '{Createdby}', '{Texto3}', '{Obra}', '{DonoObra}', '{Morada}', '{Localidade}', '{CodPostal}', '{CodPostalLocal}', '{Anexos}', '{Texto4}', '{Texto5}', '{estado}', '{atencao}')";



                ProductContext.MotorLE.DSO.ExecuteSQL(query);



                return Request.CreateResponse(HttpStatusCode.OK, "Ofício criado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao criar ofício: {ex.Message}");
            }
        }

        [Authorize]
        [Route("Eliminar/{codigo}")]
        [HttpGet]
        public HttpResponseMessage Eliminar(string codigo)
        {
            try
            {
                if (string.IsNullOrEmpty(codigo))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Código inválido.");
                }
                string query = $"UPDATE TDU_Oficios SET CDU_isactive = 0 where CDU_codigo = '{codigo}'";

                ProductContext.MotorLE.DSO.ExecuteSQL(query);

                return Request.CreateResponse(HttpStatusCode.OK, "Ofício excluído com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao excluir ofício: {ex.Message}");
            }
        }

        [Authorize]
        [Route("Atualizar")]
        [HttpPut]
        public HttpResponseMessage Atualizar([FromBody] OficioDto oficioAtualizado)
        {
            try
            {
                // Validate if the incoming data is valid
                if (oficioAtualizado == null || string.IsNullOrEmpty(oficioAtualizado.Codigo))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados inválidos.");
                }

                // Extract the updated values from the DTO

                var Codigo = oficioAtualizado.Codigo;
                var Assunto = oficioAtualizado.Assunto;
                var Data = oficioAtualizado.Data;
                var Remetente = oficioAtualizado.Remetente;
                var Email = oficioAtualizado.Email;
                var Texto1 = oficioAtualizado.Texto1;
                var Texto2 = oficioAtualizado.Texto2;
                var Template = oficioAtualizado.Template;
                var Createdby = oficioAtualizado.Createdby;
                var Texto3 = oficioAtualizado.Texto3;
                var Obra = oficioAtualizado.Obra;
                var DonoObra = oficioAtualizado.donoObra;
                var Morada = oficioAtualizado.Morada;
                var Localidade = oficioAtualizado.Localidade;
                var CodPostal = oficioAtualizado.CodPostal;
                var CodPostalLocal = oficioAtualizado.CodPostalLocal;
                var Anexos = oficioAtualizado.Anexos;
                var Texto4 = oficioAtualizado.Texto4;
                var Texto5 = oficioAtualizado.Texto5;
                var Atencao = oficioAtualizado.Atencao;
                // Prepare the SQL query to update the record
                string query = $@"UPDATE TDU_Oficios
                SET 
                    CDU_assunto = '{Assunto}', 
                    CDU_data = '{Data}', 
                    CDU_remetente = '{Remetente}', 
                    CDU_email = '{Email}', 
                    CDU_texto1 = '{Texto1}', 
                    CDU_texto2 = '{Texto2}', 
                    CDU_template = '{Template}', 
                    CDU_createdby = '{Createdby}', 
                    CDU_texto3 = '{Texto3}', 
                    CDU_obra = '{Obra}', 
                    CDU_DonoObra = '{DonoObra}', 
                    CDU_Morada = '{Morada}', 
                    CDU_Localidade = '{Localidade}', 
                    CDU_CodPostal = '{CodPostal}', 
                    CDU_CodPostalLocal = '{CodPostalLocal}', 
                    CDU_Anexos = '{Anexos}', 
                    CDU_texto4 = '{Texto4}', 
                    CDU_texto5 = '{Texto5}',
                    CDU_AC = '{Atencao}'
                WHERE CDU_codigo = '{Codigo}'";


                // Execute the update query
                ProductContext.MotorLE.DSO.ExecuteSQL(query);

                // Return success response
                return Request.CreateResponse(HttpStatusCode.OK, "Ofício atualizado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar ofício: {ex.Message}");
            }
        }

        [Authorize]
        [Route("Detalhes/{codigo}")]
        [HttpGet]
        public HttpResponseMessage Detalhes(string codigo)
        {
            try
            {
                if (string.IsNullOrEmpty(codigo))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Código inválido.");
                }

                string query = $"SELECT * FROM TDU_Oficios WHERE CDU_codigo = '{codigo}'";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Ofício não encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter detalhes: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarObras")]
        [HttpGet]
        public HttpResponseMessage ListarObras()
        {
            try
            {



                // Prepare the SQL query to update the record
                string query = $@"SELECT * FROM COP_obras 
                                WHERE (estado = 'CONS' OR estado='Conc') AND Tipo = 'O' 
                                ORDER BY descricao ASC;
                                ";

                // Execute the update query
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, response);
                }
                // Return success response
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao listar obras: {ex.Message}");
            }
        }




        [Authorize]
        [Route("ListarEntidades")]
        [HttpGet]
        public HttpResponseMessage ListarEntidades()
        {
            try
            {



                // Prepare the SQL query to update the record
                string query = $@"SELECT * FROM Geral_Entidade";

                // Execute the update query
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, response);
                }
                // Return success response
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar ofício: {ex.Message}");
            }
        }

        [Authorize]
        [Route("DaUltimoID")]
        [HttpGet]
        public HttpResponseMessage DaUltimoID()
        {
            try
            {
                // Prepare the SQL query to update the record
                string query = @"SELECT TOP 1 
                                SUBSTRING(CDU_codigo, 10, 4) AS Conteudo_Central
                                FROM TDU_Oficios
                                ORDER BY CDU_codigo DESC;";

                // Execute the update query
                var response = ProductContext.MotorLE.Consulta(query);

                if (response.NumLinhas() <= 0)
                {
                    return Request.CreateResponse(HttpStatusCode.OK, "000");
                }
                // Return success response
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar ofício: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetEntidade/{entidadeId}")]
        [HttpGet]
        public HttpResponseMessage GetEntidade(string entidadeId)
        {
            try
            {

                // Prepare the SQL query to update the record
                string query = $@"SELECT * FROM Geral_Entidade where EntidadeId = {entidadeId}";

                // Execute the update query
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, response);
                }
                // Return success response
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar ofício: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetEmail/{entidadeId}")]
        [HttpGet]
        public HttpResponseMessage GetEmail(string entidadeId)
        {
            try
            {

                // Prepare the SQL query to update the record
                string query = $@"SELECT * FROM Geral_Entidade_Contactos where EntidadeID = '{entidadeId}'";

                // Execute the update query
                var response = ProductContext.MotorLE.Consulta(query);

                if (response.NumLinhas() <= 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "response");
                }
                // Return success response
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar ofício: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetEntidadeCode/{obraCode}")]
        [HttpGet]
        public HttpResponseMessage GetEntidadeCode(string obraCode)
        {
            try
            {

                // Prepare the SQL query to update the record
                string query = $@"SELECT * FROM COP_obras where ID = '{obraCode}'";

                // Execute the update query
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, response);
                }
                // Return success response
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar ofício: {ex.Message}");
            }
        }
    }


    public class AtualizarTrabalhadorDto
    {
        public string Documento { get; set; }
        public string Validade { get; set; }
        public string Anexo { get; set; }
        public string Contribuinte { get; set; }
        public string IdEntidade { get; set; }
    }

    public class AtualizarEquipamentoDto
    {
        public string Documento { get; set; }
        public string Validade { get; set; }
        public string Anexo { get; set; }
        public string Marca { get; set; }
        public string IdEntidade { get; set; }
    }




    public class TrabalhadorDto
    {
        public string Contribuinte { get; set; }
        public string Nome { get; set; }
        public string Funcao { get; set; }
        public string NumSegurancaSocial { get; set; }
        public string DataNascimento { get; set; }
        public string IdEntidade { get; set; }

        public string Caminho1 { get; set; }
        public string Caminho2 { get; set; }
        public string Caminho3 { get; set; }
        public string Caminho4 { get; set; }
        public string Caminho5 { get; set; }

        public bool Anexo1 { get; set; }
        public bool Anexo2 { get; set; }
        public bool Anexo3 { get; set; }
        public bool Anexo4 { get; set; }
        public bool Anexo5 { get; set; }
    }

    public class EquipamentoDto
    {
        public string Tipo { get; set; }
        public string Marca { get; set; }
        public string Serie { get; set; }
        public string IdEntidade { get; set; }

        public string Caminho1 { get; set; }
        public string Caminho2 { get; set; }
        public string Caminho3 { get; set; }
        public string Caminho4 { get; set; }
        public string Caminho5 { get; set; }

        public bool Anexo1 { get; set; }
        public bool Anexo2 { get; set; }
        public bool Anexo3 { get; set; }
        public bool Anexo4 { get; set; }
        public bool Anexo5 { get; set; }

        public string CBSeguro { get; set; }
    }

    public class AprovacaoConcurso
    {
        public string Id { get; set; }
        public string Responsavel { get; set; }
    }


    [RoutePrefix("SharePoint")]
    public class SharePointController : ApiController
    {
        [Authorize]
        [Route("ListarEntidadesSGS")]
        [HttpGet]
        public HttpResponseMessage ListarEntidadesSGS()
        {
            try
            {
                var query = "SELECT id,Nome frOM Geral_Entidade WHERE CDU_TrataSGS = 1";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma entidade encontrada.");

                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter entidades: {ex.Message}");
            }
        }


        [Authorize]
        [Route("GetEntidade/{id}")]
        [HttpGet]
        public HttpResponseMessage GetEntidade(string id)
        {
            try
            {
                var query = $"SELECT * FROM Geral_Entidade WHERE id = '{id}'";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma entidade encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter entidade: {ex.Message}");
            }
        }

        [Authorize]
        [Route("UpdateEntidade/{documento}/{validade}/{Anexo}/{idEntidade}")]
        [HttpPut]
        public HttpResponseMessage UpdateEntidade(string documento, string validade, string Anexo, string idEntidade)
        {
            try
            {

                var query = $@"
                            UPDATE Geral_Entidade
                            SET {documento} = '{validade}', 
                                {Anexo} = 1
                            WHERE Id = '{idEntidade}'";

                var result = ProductContext.MotorLE.DSO.ExecuteSQL(query);

                if (result == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Entidade não encontrada ou nada foi alterado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Entidade atualizada com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar entidade: {ex.Message}");
            }
        }

        [Authorize]
        [Route("UpdateTrabalhador")]
        [HttpPut]
        public HttpResponseMessage UpdateTrabalhador([FromBody] AtualizarTrabalhadorDto dto)
        {
            try
            {
                var query = $@"
            UPDATE TDU_AD_Trabalhadores
            SET {dto.Documento} = '{dto.Validade}', 
                {dto.Anexo} = 1
            WHERE contribuinte = '{dto.Contribuinte}' AND id_empresa = '{dto.IdEntidade}'";

                var result = ProductContext.MotorLE.DSO.ExecuteSQL(query);

                if (result == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Trabalhador não encontrado ou nada foi alterado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Trabalhador atualizado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar trabalhador: {ex.Message}");
            }
        }

        [Authorize]
        [Route("InsertTrabalhador")]
        [HttpPut]
        public HttpResponseMessage InsertTrabalhador([FromBody] TrabalhadorDto dto)
        {
            try
            {

                int anexo1 = dto.Anexo1 ? 1 : 0;
                int anexo2 = dto.Anexo2 ? 1 : 0;
                int anexo3 = dto.Anexo3 ? 1 : 0;
                int anexo4 = dto.Anexo4 ? 1 : 0;
                int anexo5 = dto.Anexo5 ? 1 : 0;

                var query = $@"
        INSERT INTO TDU_AD_Trabalhadores (
            Contribuinte, Nome, categoria, seguranca_social, data_nascimento, id_empresa,
            Caminho1, Caminho2, Caminho3, Caminho4, Caminho5,
            Anexo1, Anexo2, Anexo3, Anexo4, Anexo5
        )
        VALUES (
            '{dto.Contribuinte}', '{dto.Nome}', '{dto.Funcao}', '{dto.NumSegurancaSocial}',
            '{dto.DataNascimento}', '{dto.IdEntidade}', '{dto.Caminho1}', '{dto.Caminho2}',
            '{dto.Caminho3}', '{dto.Caminho4}', '{dto.Caminho5}',
            {anexo1}, {anexo2}, {anexo3}, {anexo4}, {anexo5}
        )";
                var result = ProductContext.MotorLE.DSO.ExecuteSQL(query);

                if (result == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Trabalhador não encontrado ou nada foi alterado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Trabalhador atualizado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar trabalhador: {ex.Message}");
            }
        }

        [Authorize]
        [Route("UpdateEquipamento")]
        [HttpPut]
        public HttpResponseMessage UpdateEquipamento([FromBody] AtualizarEquipamentoDto dto)
        {
            try
            {
                var query = $@"
            UPDATE TDU_AD_Equipamentos
            SET {dto.Documento} = '{dto.Validade}', 
                {dto.Anexo} = 1
            WHERE marca = '{dto.Marca}' AND id_empresa = '{dto.IdEntidade}'";
                var result = ProductContext.MotorLE.DSO.ExecuteSQL(query);

                if (result == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Equipamento não encontrado ou nada foi alterado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Equipamento atualizado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar equipamento: {ex.Message}");
            }
        }

        [Authorize]
        [Route("InsertEquipamento")]
        [HttpPut]
        public HttpResponseMessage InsertEquipamento([FromBody] EquipamentoDto dto)
        {
            try
            {
                int anexo1 = dto.Anexo1 ? 1 : 0;
                int anexo2 = dto.Anexo2 ? 1 : 0;
                int anexo3 = dto.Anexo3 ? 1 : 0;
                int anexo4 = dto.Anexo4 ? 1 : 0;
                int anexo5 = dto.Anexo5 ? 1 : 0;


                string cSeguro = "NA";

                var query = $@"
            INSERT INTO TDU_AD_Equipamentos (
                tipo, marca, serie, id_empresa,
                caminho1, caminho2, caminho3, caminho4, caminho5,
                anexo1, anexo2, anexo3, anexo4, anexo5, cBSeguro
            )
            VALUES (
                '{dto.Tipo}', '{dto.Marca}', '{dto.Serie}', '{dto.IdEntidade}',
                '{dto.Caminho1}', '{dto.Caminho2}', '{dto.Caminho3}', '{dto.Caminho4}', '{dto.Caminho5}',
                {anexo1}, {anexo2}, {anexo3}, {anexo4}, {anexo5}, '{cSeguro}'
            )";

                var result = ProductContext.MotorLE.DSO.ExecuteSQL(query);

                if (result == 0)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Equipamento não encontrado ou nada foi alterado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Equipamento inserido com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao inserir equipamento: {ex.Message}");
            }
        }


        [Authorize]
        [Route("ListarTrabalhadores/{id_entidade}")]
        [HttpGet]
        public HttpResponseMessage ListarTrabalhadores(string id_entidade)
        {
            try
            {
                var query = $"SELECT * FROM TDU_AD_Trabalhadores WHERE id_empresa = '{id_entidade}'";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum trabalhador encontrado 2.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter trabalhadores 2: {ex.Message}");
            }
        }

        [Authorize]
        [Route("ListarEquipamentos/{id_entidade}")]
        [HttpGet]
        public HttpResponseMessage ListarEquipamentos(string id_entidade)
        {
            try
            {
                var query = $"SELECT * FROM TDU_AD_Equipamentos WHERE id_empresa = '{id_entidade}'";
                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum equipamento encontrado.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter equipamentos: {ex.Message}");
            }
        }


    }

    [RoutePrefix("Concursos")]
    public class ConcursosController : ApiController
    {
        [Authorize]
        [Route("GetListaConcursos")]
        [HttpGet]
        public HttpResponseMessage GetListaConcursos()
        {
            try
            {
                string query = $@"
SELECT 
  CP.Codigo,
  CP.Titulo,
  CP.DataProposta,
  CP.PrecoBase,
  GE.Nome,
  CF.Descricao AS FormaContrato,
  CT.Descricao AS TipoProposta,
  Z.Desig AS Zona,
  CTT.Descricao AS TipoObra,
  F.Factor,
  F.Descricao AS DescricaoCriterio,
  OC.Peso,
  CP.DataEntregaPropostas
FROM 
  COP_Obras AS CP
left JOIN Geral_Entidade AS GE ON CP.EntidadeIDA = GE.EntidadeId
LEFT JOIN COP_FormaContratos AS CF ON CP.TipoCont = CF.Tipo 
LEFT JOIN Geral_Zona AS Z ON CP.ZonaID = Z.Zona
LEFT JOIN Cop_TipoPropostas AS CT ON CP.TipoProp = CT.Tipo
LEFT JOIN Cop_TipoObras AS CTT ON CP.TipoObra = CTT.Tipo
LEFT JOIN COP_ObraCriterios AS OC ON OC.ObraID = CP.ID
LEFT JOIN COP_Factores AS F ON OC.FactorID = F.ID
WHERE 
  CP.Tipo = 'O' AND
  CP.Estado = 'CONC'  AND
  CP.CDU_Autorizacao = '1'
ORDER BY 
       CP.Codigo DESC
";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum concurso encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de concursos: {ex.Message}");
            }

        }

        [Authorize]
        [Route("AtualizaEstadoAprovado")]
        [HttpPost]
        public HttpResponseMessage AtualizaEstadoAprovado([FromBody] AprovacaoConcurso aprovacaoConcurso)
        {
            try
            {

                var queryAprovado = $@"UPDATE COP_Obras
                                    SET CDU_EstadoConcurso = 'Aprovado',
                                        CDU_ResponsavelEstadoConcurso = '{aprovacaoConcurso.Responsavel}',
                                        CDU_DataEstadoConcurso = GETDATE(),
CDU_Autorizacao = '0'
                                    WHERE Codigo = '{aprovacaoConcurso.Id}';
                                    ";
                ProductContext.MotorLE.DSO.ExecuteSQL(queryAprovado);
                return Request.CreateResponse(HttpStatusCode.OK, "Aprovado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de concursos: {ex.Message}");
            }

        }

        [Authorize]
        [Route("AtualizaEstadoRescusado")]
        [HttpPost]
        public HttpResponseMessage AtualizaEstadoRescusado([FromBody] AprovacaoConcurso aprovacaoConcurso)
        {
            try
            {

                var queryAprovado = $@"UPDATE COP_Obras
                                    SET CDU_EstadoConcurso = 'Recusado',
                                        CDU_ResponsavelEstadoConcurso = '{aprovacaoConcurso.Responsavel}',
                                        CDU_DataEstadoConcurso = GETDATE(),
CDU_Autorizacao = '0'
                                    WHERE Codigo = '{aprovacaoConcurso.Id}';
                                    ";
                ProductContext.MotorLE.DSO.ExecuteSQL(queryAprovado);
                return Request.CreateResponse(HttpStatusCode.OK, "Recusado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de concursos: {ex.Message}");
            }

        }



        [Authorize]
        [Route("GetListaFaltasFuncionario/{codFuncionario}")]
        [HttpGet]
        public HttpResponseMessage GetListaFaltasFuncionario(string codFuncionario)
        {
            try
            {
                string query = $@"Select * from CadastroFaltas where Funcionario = '{codFuncionario}' ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de falta encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de tipos de faltas: {ex.Message}");
            }

        }
    }


    [RoutePrefix("AlteracoesMensais")]
    public class AlteracoesMensaisController : ApiController
    {


        [Authorize]
        [Route("Feriados")]
        [HttpGet]
        public HttpResponseMessage Feriados()
        {
            try
            {
                string query = @"SELECT * FROM Feriados
        ";

                var response = ProductContext.MotorLE.Consulta(query);
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter Feriados: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetListaTipoFaltas")]
        [HttpGet]
        public HttpResponseMessage GetListaTipoFaltas()
        {
            try
            {
                string query = $@"Select * from Faltas where CDU_IntegraLink=1";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de falta encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de tipos de faltas: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetListaFaltasFuncionario/{codFuncionario}")]
        [HttpGet]
        public HttpResponseMessage GetListaFaltasFuncionario(string codFuncionario)
        {
            try
            {
                string query = $@"Select * from CadastroFaltas where Funcionario = '{codFuncionario}' ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de falta encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de tipos de faltas: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetListaFaltasTodosFuncionarios")]
        [HttpGet]
        public HttpResponseMessage GetListaFaltasTodosFuncionarios()
        {
            try
            {
                string query = $@"Select * from CadastroFaltas ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de falta encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de tipos de faltas: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetListaFeriasFuncionario/{codFuncionario}")]
        [HttpGet]
        public HttpResponseMessage GetListaFeriasFuncionario(string codFuncionario)
        {
            try
            {
                string query = $@"Select * from RHP_Ferias where Funcionario = '{codFuncionario}' ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Férias não encontradas.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de férias: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetListaFeriasTodosFuncionarios")]
        [HttpGet]
        public HttpResponseMessage GetListaFeriasTodosFuncionarios()
        {
            try
            {
                string query = $@"Select * from RHP_Ferias ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Férias não encontradas.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de férias: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetTotalizadorFeriasTodosFuncionarios")]
        [HttpGet]
        public HttpResponseMessage GetTotalizadorFeriasTodosFuncionarios()
        {
            try
            {
                string query = $@"select* from FuncInfFerias";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Férias não encontradas.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de férias: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetTotalizadorFeriasFuncionario/{codFuncionario}")]
        [HttpGet]
        public HttpResponseMessage GetTotalizadorFeriasTodosFuncionarios(string codFuncionario)
        {
            try
            {
                string query = $@"SELECT TOP (1) *
FROM FuncInfFerias
WHERE Funcionario =  '{codFuncionario}' ORDER BY Ano DESC;
";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Férias não encontradas.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de férias: {ex.Message}");
            }

        }



        [Authorize]
        [Route("GetHorarioFuncionario/{codFuncionario}")]
        [HttpGet]
        public HttpResponseMessage GetHorarioFuncionario(string codFuncionario)
        {
            try
            {
                string query = $@"select* from FuncHorarios where Funcionario= '{codFuncionario}'";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Horário do funcionário não encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter horário do trabalhador: {ex.Message}");
            }

        }



        [Authorize]
        [Route("GetNomeFuncionario/{codFuncionario}")]
        [HttpGet]
        public HttpResponseMessage GetNomeFuncionario(string codFuncionario)
        {
            try
            {
                string query = $@"select Nome from funcionarios where codigo= '{codFuncionario}'";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nome do funcionário não encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter nome do funcionário: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetHorariosTrabalho")]
        [HttpGet]
        public HttpResponseMessage GetHorariosTrabalho()
        {
            try
            {
                string query = $@"select* from HorariosTrabalho";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Horários não encontradas.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de horários: {ex.Message}");
            }

        }



        [Authorize]
        [Route("InserirFalta")]
        [HttpPost]
        public HttpResponseMessage InserirFalta([FromBody] CadastroFaltaModel novaFalta)
        {
            try
            {
                if (novaFalta == null)
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados inválidos.");
                }

                string query = $@"
            INSERT INTO CadastroFaltas (
                Funcionario, Data, Falta, Horas, Tempo, DescontaVenc, DescontaRem,
                ExcluiProc, ExcluiEstat, Observacoes, CalculoFalta, DescontaSubsAlim,
                DataProc, NumPeriodoProcessado, JaProcessado, InseridoBloco,
                ValorDescontado, AnoProcessado, NumProc, Origem, PlanoCurso,
                IdGDOC, CambioMBase, CambioMAlt, CotizaPeloMinimo, Acerto,
                MotivoAcerto, NumLinhaDespesa, NumRelatorioDespesa,
                FuncComplementosBaixaId, DescontaSubsTurno, SubTurnoProporcional, SubAlimProporcional
            )
            VALUES (
                '{novaFalta.Funcionario}', '{novaFalta.Data:yyyy-MM-dd HH:mm:ss}', '{novaFalta.Falta}', {novaFalta.Horas}, {novaFalta.Tempo},
                {novaFalta.DescontaVenc}, {novaFalta.DescontaRem}, {novaFalta.ExcluiProc}, {novaFalta.ExcluiEstat},
                {(string.IsNullOrEmpty(novaFalta.Observacoes) ? "NULL" : $"'{novaFalta.Observacoes}'")},
                {novaFalta.CalculoFalta}, {novaFalta.DescontaSubsAlim},
                {(novaFalta.DataProc.HasValue ? $"'{novaFalta.DataProc:yyyy-MM-dd HH:mm:ss}'" : "NULL")},
                {novaFalta.NumPeriodoProcessado}, {novaFalta.JaProcessado}, {novaFalta.InseridoBloco},
                {novaFalta.ValorDescontado}, {novaFalta.AnoProcessado}, {novaFalta.NumProc},
                {(string.IsNullOrEmpty(novaFalta.Origem) ? "NULL" : $"'{novaFalta.Origem}'")},
                {(string.IsNullOrEmpty(novaFalta.PlanoCurso) ? "NULL" : $"'{novaFalta.PlanoCurso}'")},
                {(string.IsNullOrEmpty(novaFalta.IdGDOC) ? "NULL" : $"'{novaFalta.IdGDOC}'")},
                {novaFalta.CambioMBase}, {novaFalta.CambioMAlt}, {novaFalta.CotizaPeloMinimo}, {novaFalta.Acerto},
                {(string.IsNullOrEmpty(novaFalta.MotivoAcerto) ? "NULL" : $"'{novaFalta.MotivoAcerto}'")},
                {(novaFalta.NumLinhaDespesa.HasValue ? novaFalta.NumLinhaDespesa.ToString() : "NULL")},
                {(novaFalta.NumRelatorioDespesa.HasValue ? novaFalta.NumRelatorioDespesa.ToString() : "NULL")},
                {(novaFalta.FuncComplementosBaixaId.HasValue ? novaFalta.FuncComplementosBaixaId.ToString() : "NULL")},
                {novaFalta.DescontaSubsTurno}, {novaFalta.SubTurnoProporcional}, {novaFalta.SubAlimProporcional}
            )";

                var resultado = ProductContext.MotorLE.DSO.ExecuteSQL(query);
                return Request.CreateResponse(HttpStatusCode.OK, "Falta inserida com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao inserir falta: {ex.Message}");
            }
        }

        [Authorize]
        [Route("InserirFeriasFuncionario")]
        [HttpPost]
        public HttpResponseMessage InserirFeriasFuncionario([FromBody] FeriasFuncionarioModel ferias)
        {
            try
            {
                if (ferias == null)
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados inválidos.");
                }

                var ano = ferias.DataFeria.Year;


                string query1 = $@"
        INSERT INTO RHP_Ferias (
            Ano, Funcionario, DataFeria, EstadoGozo, OriginouFalta,
            TipoMarcacao, OriginouFaltaSubAlim, Duracao, Acerto, NumProc, Origem
        )
        VALUES (
            {ano},
            '{ferias.Funcionario}',
            '{ferias.DataFeria:yyyy-MM-dd HH:mm:ss}',
            {ferias.EstadoGozo},
            {ferias.OriginouFalta},
            {ferias.TipoMarcacao},
            {ferias.OriginouFaltaSubAlim},
            {ferias.Duracao.ToString(System.Globalization.CultureInfo.InvariantCulture)},
            {ferias.Acerto},
            {(ferias.NumProc.HasValue ? ferias.NumProc.ToString() : "NULL")},
            {ferias.Origem}
        )";

                ProductContext.MotorLE.DSO.ExecuteSQL(query1);

                return Request.CreateResponse(HttpStatusCode.OK, "Férias inseridas com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao inserir férias: {ex.Message}");
            }
        }



        // Novo editar ferias funcionario
        [Authorize]
        [Route("EditarFeriasFuncionario")]
        [HttpPut]
        public HttpResponseMessage EditarFeriasFuncionario([FromBody] FeriasFuncionarioModel ferias)
        {
            try
            {
                if (ferias == null)
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados inválidos.");
                }

                var ano = ferias.DataFeria.Year;

                string queryUpdate = $@"
UPDATE RHP_Ferias
SET
    Ano = {ano},
    EstadoGozo = {ferias.EstadoGozo},
    OriginouFalta = {ferias.OriginouFalta},
    TipoMarcacao = {ferias.TipoMarcacao},
    OriginouFaltaSubAlim = {ferias.OriginouFaltaSubAlim},
    Duracao = {ferias.Duracao.ToString(System.Globalization.CultureInfo.InvariantCulture)},
    Acerto = {ferias.Acerto},
    NumProc = {(ferias.NumProc.HasValue ? ferias.NumProc.ToString() : "NULL")},
    Origem = {ferias.Origem}
WHERE
    RTRIM(Funcionario) = '{ferias.Funcionario.Trim()}'
    AND CONVERT(VARCHAR(10), DataFeria, 120) = '{ferias.DataFeria:yyyy-MM-dd}'
";

                System.Diagnostics.Debug.WriteLine("QUERY FINAL: " + queryUpdate);


                int linhasAfetadas = ProductContext.MotorLE.DSO.ExecuteSQL(queryUpdate);

                if (linhasAfetadas > 0)
                    return Request.CreateResponse(HttpStatusCode.OK, "Férias atualizadas com sucesso.");
                else
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Registo de férias não encontrado.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao editar férias: {ex.Message}");
            }
        }






        [Authorize]
        [Route("EditarFalta")]
        [HttpPut]
        public HttpResponseMessage EditarFalta([FromBody] CadastroFaltaModel faltaEditada)
        {
            try
            {
                if (faltaEditada == null)
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Dados inválidos.");
                }

                // Verificação mínima para evitar editar sem saber qual a linha
                if (string.IsNullOrEmpty(faltaEditada.Funcionario) || faltaEditada.Data == default || string.IsNullOrEmpty(faltaEditada.Falta))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Chave primária incompleta (Funcionario, Data, Falta).");
                }

                string query = $@"
            UPDATE CadastroFaltas SET
                Horas = {faltaEditada.Horas},
                Tempo = {faltaEditada.Tempo},
                DescontaVenc = {faltaEditada.DescontaVenc},
                DescontaRem = {faltaEditada.DescontaRem},
                ExcluiProc = {faltaEditada.ExcluiProc},
                ExcluiEstat = {faltaEditada.ExcluiEstat},
                Observacoes = {(string.IsNullOrEmpty(faltaEditada.Observacoes) ? "NULL" : $"'{faltaEditada.Observacoes}'")},
                CalculoFalta = {faltaEditada.CalculoFalta},
                DescontaSubsAlim = {faltaEditada.DescontaSubsAlim},
                DataProc = {(faltaEditada.DataProc.HasValue ? $"'{faltaEditada.DataProc:yyyy-MM-dd HH:mm:ss}'" : "NULL")},
                NumPeriodoProcessado = {faltaEditada.NumPeriodoProcessado},
                JaProcessado = {faltaEditada.JaProcessado},
                InseridoBloco = {faltaEditada.InseridoBloco},
                ValorDescontado = {faltaEditada.ValorDescontado},
                AnoProcessado = {faltaEditada.AnoProcessado},
                NumProc = {faltaEditada.NumProc},
                Origem = {(string.IsNullOrEmpty(faltaEditada.Origem) ? "NULL" : $"'{faltaEditada.Origem}'")},
                PlanoCurso = {(string.IsNullOrEmpty(faltaEditada.PlanoCurso) ? "NULL" : $"'{faltaEditada.PlanoCurso}'")},
                IdGDOC = {(string.IsNullOrEmpty(faltaEditada.IdGDOC) ? "NULL" : $"'{faltaEditada.IdGDOC}'")},
                CambioMBase = {faltaEditada.CambioMBase},
                CambioMAlt = {faltaEditada.CambioMAlt},
                CotizaPeloMinimo = {faltaEditada.CotizaPeloMinimo},
                Acerto = {faltaEditada.Acerto},
                MotivoAcerto = {(string.IsNullOrEmpty(faltaEditada.MotivoAcerto) ? "NULL" : $"'{faltaEditada.MotivoAcerto}'")},
                NumLinhaDespesa = {(faltaEditada.NumLinhaDespesa.HasValue ? faltaEditada.NumLinhaDespesa.ToString() : "NULL")},
                NumRelatorioDespesa = {(faltaEditada.NumRelatorioDespesa.HasValue ? faltaEditada.NumRelatorioDespesa.ToString() : "NULL")},
                FuncComplementosBaixaId = {(faltaEditada.FuncComplementosBaixaId.HasValue ? faltaEditada.FuncComplementosBaixaId.ToString() : "NULL")},
                DescontaSubsTurno = {faltaEditada.DescontaSubsTurno},
                SubTurnoProporcional = {faltaEditada.SubTurnoProporcional},
                SubAlimProporcional = {faltaEditada.SubAlimProporcional}
            WHERE
                Funcionario = '{faltaEditada.Funcionario}' AND
                Data = '{faltaEditada.Data:yyyy-MM-dd HH:mm:ss}' AND
                Falta = '{faltaEditada.Falta}'
        ";

                var resultado = ProductContext.MotorLE.DSO.ExecuteSQL(query);
                return Request.CreateResponse(HttpStatusCode.OK, "Falta editada com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao editar falta: {ex.Message}");
            }
        }




        [Authorize]
        [Route("EliminarFalta/{codFuncionario}/{dataFalta}/{tipoFalta}")]
        [HttpDelete]
        public HttpResponseMessage EliminarFalta(string codFuncionario, string dataFalta, string tipoFalta)
        {
            try
            {
                DateTime data;
                if (!DateTime.TryParse(dataFalta, out data))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Data inválida.");
                }

                string query = $@"
            DELETE FROM CadastroFaltas
            WHERE RTRIM(Funcionario) = '{codFuncionario.Trim()}'
            AND CONVERT(VARCHAR(10), Data, 120) = '{data:yyyy-MM-dd}'
            AND RTRIM(Falta) = '{tipoFalta.Trim()}'";

                int linhasAfetadas = ProductContext.MotorLE.DSO.ExecuteSQL(query);

                if (linhasAfetadas > 0)
                    return Request.CreateResponse(HttpStatusCode.OK, "Falta eliminada com sucesso.");
                else
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Registo de falta não encontrado.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao eliminar falta: {ex.Message}");
            }
        }




        [Authorize]
        [Route("EliminarFeriasFuncionario/{codFuncionario}/{dataFeria}")]
        [HttpDelete]
        public HttpResponseMessage EliminarFeriasFuncionario(string codFuncionario, string dataFeria)
        {
            try
            {
                DateTime data;
                if (!DateTime.TryParse(dataFeria, out data))
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Data inválida.");
                }

                string query = $@"
            DELETE FROM RHP_Ferias
            WHERE RTRIM(Funcionario) = '{codFuncionario.Trim()}'
            AND CONVERT(VARCHAR(10), DataFeria, 120) = '{data:yyyy-MM-dd}'";

                int linhasAfetadas = ProductContext.MotorLE.DSO.ExecuteSQL(query);

                if (linhasAfetadas > 0)
                    return Request.CreateResponse(HttpStatusCode.OK, "Férias eliminadas com sucesso.");
                else
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Registo de férias não encontrado.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao eliminar férias: {ex.Message}");
            }
        }





        // PARTES DIARIAS

        [Authorize]
        [Route("GetPartesDiarias/{obraID}")]
        [HttpGet]
        public HttpResponseMessage GetPartesDiarias(string obraID, string dataInicio = null, string dataFim = null)
        {
            try
            {
                // Monta o filtro base de ObraID - já protegido por parâmetro, mas aqui vai na string, então usa aspas simples
                string filtro = $"WHERE ObraID = '{obraID.Replace("'", "''")}'";

                // Se as datas forem passadas, adiciona filtro
                if (!string.IsNullOrEmpty(dataInicio) && !string.IsNullOrEmpty(dataFim))
                {
                    // Para evitar SQL Injection, faz uma validação simples das datas
                    DateTime dtInicio, dtFim;
                    if (DateTime.TryParse(dataInicio, out dtInicio) && DateTime.TryParse(dataFim, out dtFim))
                    {
                        filtro += $" AND FP.Data BETWEEN '{dtInicio:yyyy-MM-dd}' AND '{dtFim:yyyy-MM-dd}'";
                    }
                    else
                    {
                        return Request.CreateResponse(HttpStatusCode.BadRequest, "Datas inválidas.");
                    }
                }

                string query = $@"
            SELECT FP.*, FPI.*
            FROM COP_FichasPessoal AS FP
            INNER JOIN COP_FichasPessoalItems AS FPI ON FP.ID = FPI.FichasPessoalID
            {filtro}
            ORDER BY Numero DESC";

                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null || response.Vazia())
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma parte diária encontrada.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter partes diárias: {ex.Message}");
            }
        }



        [Authorize]
        [Route("GetPartesDiariasColaborador/{obraID}")]
        [HttpGet]
        public HttpResponseMessage GetPartesDiariasColaborador(string obraID, string dataInicio = null, string dataFim = null, string colaboradorID = null)
        {
            try
            {
                // Monta o filtro base de ObraID - já protegido por parâmetro, mas aqui vai na string, então usa aspas simples
                string filtro = $"WHERE FP.ObraID = '{obraID.Replace("'", "''")}'";

                // Filtro por data (se fornecido)
                if (!string.IsNullOrEmpty(dataInicio) && !string.IsNullOrEmpty(dataFim))
                {
                    if (DateTime.TryParse(dataInicio, out DateTime dtInicio) && DateTime.TryParse(dataFim, out DateTime dtFim))
                    {
                        filtro += $" AND FP.Data BETWEEN '{dtInicio:yyyy-MM-dd}' AND '{dtFim:yyyy-MM-dd}'";
                    }
                    else
                    {
                        return Request.CreateResponse(HttpStatusCode.BadRequest, "Datas inválidas.");
                    }
                }

                // Filtro por colaborador (se fornecido)
                // Filtro por colaborador (se fornecido)
                if (!string.IsNullOrEmpty(colaboradorID))
                {
                    colaboradorID = colaboradorID.Trim();

                    if (int.TryParse(colaboradorID, out int colaboradorIDInt))
                    {
                        filtro += $" AND FPI.ColaboradorID = {colaboradorIDInt}";
                    }
                    else
                    {
                        return Request.CreateResponse(HttpStatusCode.BadRequest, "ColaboradorID inválido.");
                    }
                }


                string query = $@"
            SELECT FP.*, FPI.*
            FROM COP_FichasPessoal AS FP
            INNER JOIN COP_FichasPessoalItems AS FPI ON FP.ID = FPI.FichasPessoalID
            {filtro}
            ORDER BY FP.Numero DESC";

                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null || response.Vazia())
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma parte diária encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter partes diárias: {ex.Message}");
            }
        }

        [Authorize]
        [Route("GetPartesDiariasEncarregado/{obraID}")]
        [HttpGet]
        public HttpResponseMessage GetPartesDiariasEncarregado(string obraID, string dataInicio = null, string dataFim = null, string colaboradorID = null)
        {
            try
            {
                // Monta o filtro base de ObraID - já protegido por parâmetro, mas aqui vai na string, então usa aspas simples
                string filtro = $"WHERE FP.ObraID = '{obraID.Replace("'", "''")}'";

                // Filtro por data (se fornecido)
                if (!string.IsNullOrEmpty(dataInicio) && !string.IsNullOrEmpty(dataFim))
                {
                    if (DateTime.TryParse(dataInicio, out DateTime dtInicio) && DateTime.TryParse(dataFim, out DateTime dtFim))
                    {
                        filtro += $" AND FP.Data BETWEEN '{dtInicio:yyyy-MM-dd}' AND '{dtFim:yyyy-MM-dd}'";
                    }
                    else
                    {
                        return Request.CreateResponse(HttpStatusCode.BadRequest, "Datas inválidas.");
                    }
                }

                // Filtro por colaborador (se fornecido)
                // Filtro por colaborador (se fornecido)
                if (!string.IsNullOrEmpty(colaboradorID))
                {
                    colaboradorID = colaboradorID.Trim();

                    if (int.TryParse(colaboradorID, out int colaboradorIDInt))
                    {
                        filtro += $" AND FP.ColaboradorID = {colaboradorIDInt}";
                    }
                    else
                    {
                        return Request.CreateResponse(HttpStatusCode.BadRequest, "ColaboradorID inválido.");
                    }
                }


                string query = $@"
            SELECT FP.*, FPI.*
            FROM COP_FichasPessoal AS FP
            INNER JOIN COP_FichasPessoalItems AS FPI ON FP.ID = FPI.FichasPessoalID
            {filtro}
            ORDER BY FP.Numero DESC";

                var response = ProductContext.MotorLE.Consulta(query);

                if (response == null || response.Vazia())
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhuma parte diária encontrada.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, response);
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter partes diárias: {ex.Message}");
            }
        }


        [Authorize]
        [Route("DeleteParteDiaria")]
        [HttpDelete]
        public HttpResponseMessage DeleteParteDiaria([FromUri] int numero, [FromUri] string obraID)
        {
            try
            {
                // Sanitizar obraID para evitar SQL injection
                obraID = obraID.Replace("'", "''");

                // Buscar ID da ficha pela combinação Numero + ObraID
                var queryBuscarID = $@"
            SELECT ID FROM COP_FichasPessoal 
            WHERE Numero = {numero} AND ObraID = '{obraID}'";

                var listaID = ProductContext.MotorLE.Consulta(queryBuscarID);

                if (listaID.Vazia())
                    return Request.CreateResponse(HttpStatusCode.NotFound, $"Parte diária número {numero} da obra {obraID} não encontrada.");

                listaID.Inicio();
                var fichaID = (Guid)listaID.Valor("ID");
                listaID.Termina();

                // Deletar itens primeiro (por FK)
                var queryDeleteItens = $@"DELETE FROM COP_FichasPessoalItems WHERE FichasPessoalID = '{fichaID}'";
                ProductContext.MotorLE.DSO.ExecuteSQL(queryDeleteItens);

                // Deletar cabeçalho
                var queryDeleteCabecalho = $@"DELETE FROM COP_FichasPessoal WHERE ID = '{fichaID}'";
                ProductContext.MotorLE.DSO.ExecuteSQL(queryDeleteCabecalho);

                return Request.CreateResponse(HttpStatusCode.OK, $"Parte diária número {numero} da obra {obraID} eliminada com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao eliminar parte diária: {ex.Message}");
            }
        }

        [Authorize]
        [Route("DeleteParteDiariaItem")]
        [HttpDelete]
        public HttpResponseMessage DeleteParteDiariaItem([FromUri] Guid itemID)
        {
            try
            {
                // Verifica se o item existe
                var queryCheck = $@"
            SELECT ID FROM COP_FichasPessoalItems 
            WHERE ID = '{itemID}'";

                var lista = ProductContext.MotorLE.Consulta(queryCheck);

                if (lista.Vazia())
                    return Request.CreateResponse(HttpStatusCode.NotFound, $"Item com ID {itemID} não encontrado.");

                // Apaga o item
                var queryDelete = $@"
            DELETE FROM COP_FichasPessoalItems 
            WHERE ID = '{itemID}'";

                ProductContext.MotorLE.DSO.ExecuteSQL(queryDelete);

                return Request.CreateResponse(HttpStatusCode.OK, $"Item {itemID} eliminado com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao eliminar item: {ex.Message}");
            }
        }


        [Authorize]
        [Route("InsertParteDiariaItem")]
        [HttpPut]
        public HttpResponseMessage InsertParteDiariaItem([FromBody] ParteDiariaRequestDto request)
        {
            try
            {
                int novoNumero = 1;

                var queryMaxNumero = "SELECT MAX(Numero) AS UltimoNumero FROM COP_FichasPessoal";
                var listaMax = ProductContext.MotorLE.Consulta(queryMaxNumero);

                if (!listaMax.Vazia())
                {
                    listaMax.Inicio();

                    if (!listaMax.NoFim() && listaMax.Valor("UltimoNumero") != null && listaMax.Valor("UltimoNumero").ToString() != "")
                    {
                        novoNumero = Convert.ToInt32(listaMax.Valor("UltimoNumero")) + 1;
                    }

                    listaMax.Termina();
                }



                var fichaID = Guid.NewGuid();
                var cab = request.Cabecalho;

                // 2. Inserir Cabeçalho
                var queryCabecalho = $@"
            INSERT INTO COP_FichasPessoal (
                ID, Numero, ObraID, Data, Notas,
                CabecMovCBLID, LigaCBL, CriadoPor, Utilizador, DataUltimaActualizacao,
                DocumentoID, TipoEntidade, SubEmpreiteiroID, ResponsavelID, Validado, ColaboradorID
            ) VALUES (
                '{fichaID}',
                {novoNumero},
                '{cab.ObraID}',
                '{cab.Data:yyyy-MM-dd}',
                '{cab.Notas?.Replace("'", "''")}',
                NULL,
                -1,
                '{cab.CriadoPor}',
                '{cab.Utilizador}',
                GETDATE(),
                '{cab.DocumentoID}',
                '{cab.TipoEntidade}',
                NULL,
                NULL,
                0,
                {(cab.ColaboradorID.HasValue ? cab.ColaboradorID.Value.ToString() : "NULL")}
            )
        ";

                ProductContext.MotorLE.DSO.ExecuteSQL(queryCabecalho);

                // 3. Inserir Itens
                foreach (var item in request.Itens)
                {
                    // Buscar o preco unitario do colaborador no banco usando IDOperador (ColaboradorID)
                    var queryPreco = $@"
        SELECT F.CustoPadrao 
        FROM GPR_Operadores O
        INNER JOIN Funcionarios F ON F.Codigo = O.Funcionario
        WHERE O.IDOperador = {item.ColaboradorID}
    ";

                    var listaPreco = ProductContext.MotorLE.Consulta(queryPreco);

                    decimal precoUnitario = 0;
                    if (!listaPreco.Vazia())
                    {
                        listaPreco.Inicio();
                        if (!listaPreco.NoFim() && listaPreco.Valor("CustoPadrao") != null)
                        {
                            precoUnitario = Convert.ToDecimal(listaPreco.Valor("CustoPadrao"));
                        }
                        listaPreco.Termina();
                    }

                    var itemID = Guid.NewGuid();


                    var queryItem = "";
                    if (string.IsNullOrEmpty(item.TipoHoraID))
                    {
                        queryItem = $@"
        INSERT INTO COP_FichasPessoalItems (
            ID, FichasPessoalID, ComponenteID, Funcionario, ClasseID,
            SubEmpID, NumHoras, PrecoUnit, TipoEntidade, ColaboradorID,
            Data, TotalHoras, Integrado
        ) VALUES (
            '{itemID}',
            '{fichaID}',
            {item.ComponenteID},
            '{item.Funcionario.Replace("'", "''")}',
            {item.ClasseID},
            {(item.SubEmpID.HasValue ? item.SubEmpID.Value.ToString() : "NULL")},
            {item.NumHoras.ToString(CultureInfo.InvariantCulture)},
            {precoUnitario.ToString(CultureInfo.InvariantCulture)},
            '{item.TipoEntidade}',
            {item.ColaboradorID},
            '{item.Data:yyyy-MM-dd}',
            0,
            0
        )
    ";
                    }
                    else
                    {
                        queryItem = $@"
        INSERT INTO COP_FichasPessoalItems (
            ID, FichasPessoalID, ComponenteID, Funcionario, ClasseID,
            SubEmpID, NumHoras, PrecoUnit, TipoEntidade, ColaboradorID,
            Data, TotalHoras, Integrado, TipoHoraID
        ) VALUES (
            '{itemID}',
            '{fichaID}',
            {item.ComponenteID},
            '{item.Funcionario.Replace("'", "''")}',
            {item.ClasseID},
            {(item.SubEmpID.HasValue ? item.SubEmpID.Value.ToString() : "NULL")},
            {item.NumHoras.ToString(CultureInfo.InvariantCulture)},
            {precoUnitario.ToString(CultureInfo.InvariantCulture)},
            '{item.TipoEntidade}',
            {item.ColaboradorID},
            '{item.Data:yyyy-MM-dd}',
            0,
            0,
            '{item.TipoHoraID}'
        )
    ";
                    }



                    ProductContext.MotorLE.DSO.ExecuteSQL(queryItem);
                }


                return Request.CreateResponse(HttpStatusCode.OK, $"Parte diária inserida com sucesso. Número gerado: {novoNumero}");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao inserir parte diária: {ex.Message}");
            }
        }



        [Authorize]
        [Route("InsertParteDiariaEquipamento")]
        [HttpPut]
        public HttpResponseMessage InsertParteDiariaEquipamento([FromBody] ParteDiariaEquipamentoRequestDto request)
        {
            try
            {
                // 1) Descobrir próximo Número
                int novoNumero = 1;
                var queryMaxNumero = "SELECT MAX(Numero) AS UltimoNumero FROM COP_FichasEquipamento";
                var listaMax = ProductContext.MotorLE.Consulta(queryMaxNumero);

                if (!listaMax.Vazia())
                {
                    listaMax.Inicio();
                    if (!listaMax.NoFim() && listaMax.Valor("UltimoNumero") != null && listaMax.Valor("UltimoNumero").ToString() != "")
                        novoNumero = Convert.ToInt32(listaMax.Valor("UltimoNumero")) + 1;
                    listaMax.Termina();
                }

                // 2) Inserir Cabeçalho
                var fichaID = Guid.NewGuid();
                var cab = request.Cabecalho;

                var encarregadoSql = string.IsNullOrWhiteSpace(cab.Encarregado)
                    ? "NULL"
                    : $"'{cab.Encarregado.Replace("'", "''")}'";

                var queryCabecalho = $@"
INSERT INTO COP_FichasEquipamento (
    ID, Numero, ObraID, Data, Encarregado, Notas,
    CabecMovCBLID, LigaCBL, CriadoPor, Utilizador, DataUltimaActualizacao, DocumentoID
) VALUES (
    '{fichaID}',
    {novoNumero},
    '{cab.ObraID}',
    '{cab.Data:yyyy-MM-dd}',
    {encarregadoSql},
    '{(cab.Notas ?? string.Empty).Replace("'", "''")}',
    NULL,
    -1,
    '{cab.CriadoPor}',
    '{cab.Utilizador}',
    GETDATE(),
    '{cab.DocumentoID}'
)";
                ProductContext.MotorLE.DSO.ExecuteSQL(queryCabecalho);

                // 3) Inserir Itens
                foreach (var item in request.Itens)
                {
                    // Validação opcional de SubEmpID para evitar FK
                    if (item.SubEmpID.HasValue)
                    {
                        var validaSubEmp = ProductContext.MotorLE.Consulta(
                            $"SELECT 1 FROM Geral_SubEmpreitada WHERE SubEmpId = {item.SubEmpID.Value}");
                        var existeSubEmp = !validaSubEmp.Vazia();
                        validaSubEmp.Termina();

                        if (!existeSubEmp)
                        {
                            return Request.CreateResponse(
                                HttpStatusCode.BadRequest,
                                $"SubEmpID {item.SubEmpID.Value} inexistente em Geral_SubEmpreitada (evitar erro de FK).");
                        }
                    }

                    var itemID = Guid.NewGuid();
                    var funcionarioSql = string.IsNullOrWhiteSpace(item.Funcionario)
                        ? "NULL"
                        : $"'{item.Funcionario.Replace("'", "''")}'";

                    var fornecedorSql = item.Fornecedor.HasValue ? item.Fornecedor.Value.ToString() : "NULL";
                    var subEmpSql = item.SubEmpID.HasValue ? item.SubEmpID.Value.ToString() : "NULL";
                    var horasOrdemSql = item.NumHorasOrdem.HasValue ? item.NumHorasOrdem.Value.ToString(CultureInfo.InvariantCulture) : "0";
                    var horasAvSql = item.NumHorasAvariada.HasValue ? item.NumHorasAvariada.Value.ToString(CultureInfo.InvariantCulture) : "0";
                    var precoUnit = item.PrecoUnit.HasValue ? item.PrecoUnit.Value : 0m;
                    var itemIdSql = item.ItemId.HasValue ? $"'{item.ItemId.Value}'" : "NULL";

                    var queryItem = $@"
INSERT INTO COP_FichasEquipamentoItems (
    ID, FichasEquipamentoID, ComponenteID, Funcionario, ClasseID,
    Fornecedor, SubEmpID, NumHorasTrabalho, NumHorasOrdem, NumHorasAvariada,
    PrecoUnit, ItemId
) VALUES (
    '{itemID}',
    '{fichaID}',
    {item.ComponenteID},
    {funcionarioSql},
    {item.ClasseID},
    {fornecedorSql},
    {subEmpSql},
    {item.NumHorasTrabalho.ToString(CultureInfo.InvariantCulture)},
    {horasOrdemSql},
    {horasAvSql},
    {precoUnit.ToString(CultureInfo.InvariantCulture)},
    {itemIdSql}
)";
                    ProductContext.MotorLE.DSO.ExecuteSQL(queryItem);
                }

                return Request.CreateResponse(HttpStatusCode.OK, $"Parte diária de equipamentos inserida com sucesso. Número gerado: {novoNumero}");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao inserir parte diária de equipamentos: {ex.Message}");
            }
        }






        [Authorize]
        [Route("UpdateParteDiariaItem")]
        [HttpPut]
        public HttpResponseMessage UpdateParteDiariaItem([FromBody] ParteDiariaRequestDto request)
        {
            try
            {
                var cab = request.Cabecalho;

                // Buscar o ID pelo Numero
                var queryBuscarID = $@"
            SELECT ID FROM COP_FichasPessoal WHERE Numero = {cab.Numero}
        ";

                var listaID = ProductContext.MotorLE.Consulta(queryBuscarID);

                if (listaID.Vazia())
                    return Request.CreateResponse(HttpStatusCode.NotFound, $"Parte diária com Número {cab.Numero} não encontrada.");

                listaID.Inicio();
                var fichaID = (Guid)listaID.Valor("ID");
                listaID.Termina();

                // Atualizar cabeçalho
                var queryUpdateCab = $@"
            UPDATE COP_FichasPessoal
            SET ObraID = '{cab.ObraID}',
                Data = '{cab.Data:yyyy-MM-dd}',
                Notas = '{cab.Notas?.Replace("'", "''")}',
                CriadoPor = '{cab.CriadoPor}',
                Utilizador = '{cab.Utilizador}',
                DataUltimaActualizacao = GETDATE(),
                DocumentoID = '{cab.DocumentoID}',
                TipoEntidade = '{cab.TipoEntidade}',
                ColaboradorID = {(cab.ColaboradorID.HasValue ? cab.ColaboradorID.Value.ToString() : "NULL")}
            WHERE ID = '{fichaID}'
        ";

                ProductContext.MotorLE.DSO.ExecuteSQL(queryUpdateCab);

                // Excluir itens antigos
                var queryDeleteItems = $@"DELETE FROM COP_FichasPessoalItems WHERE FichasPessoalID = '{fichaID}'";
                ProductContext.MotorLE.DSO.ExecuteSQL(queryDeleteItems);

                // Inserir novos itens (mesmo código que você já tem)
                foreach (var item in request.Itens)
                {
                    var queryPreco = $@"
                SELECT F.CustoPadrao 
                FROM GPR_Operadores O
                INNER JOIN Funcionarios F ON F.Codigo = O.Funcionario
                WHERE O.IDOperador = {item.ColaboradorID}
            ";

                    var listaPreco = ProductContext.MotorLE.Consulta(queryPreco);

                    decimal precoUnitario = 0;
                    if (!listaPreco.Vazia())
                    {
                        listaPreco.Inicio();
                        if (!listaPreco.NoFim() && listaPreco.Valor("CustoPadrao") != null)
                        {
                            precoUnitario = Convert.ToDecimal(listaPreco.Valor("CustoPadrao"));
                        }
                        listaPreco.Termina();
                    }

                    var itemID = Guid.NewGuid();

                    var queryItem = $@"
                INSERT INTO COP_FichasPessoalItems (
                    ID, FichasPessoalID, ComponenteID, Funcionario, ClasseID,
                    SubEmpID, NumHoras, PrecoUnit, TipoEntidade, ColaboradorID,
                    Data, TotalHoras, Integrado
                ) VALUES (
                    '{itemID}',
                    '{fichaID}',
                    {item.ComponenteID},
                    '{item.Funcionario.Replace("'", "''")}',
                    {item.ClasseID},
                    {(item.SubEmpID.HasValue ? item.SubEmpID.Value.ToString() : "NULL")},
                    {item.NumHoras.ToString(CultureInfo.InvariantCulture)},
                    {precoUnitario.ToString(CultureInfo.InvariantCulture)},
                    '{item.TipoEntidade}',
                    {item.ColaboradorID},
                    '{item.Data:yyyy-MM-dd}',
                    0,
                    0
                )
            ";

                    ProductContext.MotorLE.DSO.ExecuteSQL(queryItem);
                }

                return Request.CreateResponse(HttpStatusCode.OK, $"Parte diária número {cab.Numero} atualizada com sucesso.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao atualizar parte diária: {ex.Message}");
            }
        }


        [Authorize]
        [Route("GetListaClasses")]
        [HttpGet]
        public HttpResponseMessage GetListaClasses()
        {
            try
            {
                string query = $@"select * from Geral_Classe ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de Geral_Classe encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de Geral_Classe: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetListaEspecialidades")]
        [HttpGet]
        public HttpResponseMessage GetListaEspecialidades()
        {
            try
            {
                string query = $@"select * from Geral_SubEmpreitada ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de Geral_SubEmpreitada encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de Geral_SubEmpreitada: {ex.Message}");
            }

        }


        [Authorize]
        [Route("GetListaEquipamentos")]
        [HttpGet]
        public HttpResponseMessage GetListaEquipamentos()
        {
            try
            {
                string query = $@"SELECT ComponenteID, Codigo, Desig FROM Precos_Componente WHERE Tipo= 2 AND TipoValor = 0 ORDER BY 3 ";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Nenhum tipo de Precos_Componente encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter lista de Precos_Componente: {ex.Message}");
            }

        }

        [Authorize]
        [Route("GetColaboradorId/{codFuncionario}")]
        [HttpGet]
        public HttpResponseMessage GetColaboradorId(string codFuncionario)
        {
            try
            {
                string query = $@"select IDOperador from GPR_Operadores where Operador= '{codFuncionario}'";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Colaborador não encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter Colaborador: {ex.Message}");
            }

        }


        [Authorize]
        [Route("GetObraId/{codObra}")]
        [HttpGet]
        public HttpResponseMessage GetObraId(string codObra)
        {
            try
            {
                string query = $@"select Id from COP_Obras where Codigo= '{codObra}'";
                var response = ProductContext.MotorLE.Consulta(query);
                if (response == null)
                {
                    return Request.CreateResponse(HttpStatusCode.NotFound, "Obra não encontrado.");
                }
                return Request.CreateResponse(HttpStatusCode.OK, response);


            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Erro ao obter Obra: {ex.Message}");
            }

        }




    }
    public class CadastroFaltaModel
    {
        public string Funcionario { get; set; }
        public DateTime Data { get; set; }
        public string Falta { get; set; }
        public int Horas { get; set; }
        public int Tempo { get; set; }
        public int DescontaVenc { get; set; }
        public int DescontaRem { get; set; }
        public int ExcluiProc { get; set; }
        public int ExcluiEstat { get; set; }
        public string Observacoes { get; set; }
        public int CalculoFalta { get; set; }
        public int DescontaSubsAlim { get; set; }
        public DateTime? DataProc { get; set; }
        public int NumPeriodoProcessado { get; set; }
        public int JaProcessado { get; set; }
        public int InseridoBloco { get; set; }
        public decimal ValorDescontado { get; set; }
        public int AnoProcessado { get; set; }
        public int NumProc { get; set; }
        public string Origem { get; set; }
        public string PlanoCurso { get; set; }
        public string IdGDOC { get; set; }
        public decimal CambioMBase { get; set; }
        public decimal CambioMAlt { get; set; }
        public int CotizaPeloMinimo { get; set; }
        public int Acerto { get; set; }
        public string MotivoAcerto { get; set; }
        public int? NumLinhaDespesa { get; set; }
        public int? NumRelatorioDespesa { get; set; }
        public int? FuncComplementosBaixaId { get; set; }
        public int DescontaSubsTurno { get; set; }
        public int SubTurnoProporcional { get; set; }
        public int SubAlimProporcional { get; set; }
    }

    public class ParteDiariaRequestDto
    {
        public ParteDiariaCabecalhoDto Cabecalho { get; set; }
        public List<ParteDiariaItemDto> Itens { get; set; }
    }

    public class ParteDiariaCabecalhoDto
    {
        public int Numero { get; set; }
        public Guid ObraID { get; set; }
        public DateTime Data { get; set; }
        public string Notas { get; set; }
        public string CriadoPor { get; set; }
        public string Utilizador { get; set; }
        public Guid DocumentoID { get; set; }
        public string TipoEntidade { get; set; } = "O";
        public int? ColaboradorID { get; set; }
    }

    public class ParteDiariaItemDto
    {
        public int ComponenteID { get; set; }
        public string Funcionario { get; set; }
        public string TipoHoraID { get; set; }
        public int ClasseID { get; set; }
        public int? SubEmpID { get; set; }
        public decimal NumHoras { get; set; }
        public decimal PrecoUnit { get; set; }
        public string TipoEntidade { get; set; } = "O";
        public int ColaboradorID { get; set; }
        public DateTime Data { get; set; }
    }

    public class ParteDiariaEquipamentoCabecalhoDto
    {
        public Guid ObraID { get; set; }
        public DateTime Data { get; set; }
        public string Notas { get; set; }
        public string CriadoPor { get; set; }
        public string Utilizador { get; set; }
        public Guid DocumentoID { get; set; }
        public string Encarregado { get; set; } // opcional
    }

    public class ParteDiariaEquipamentoItemDto
    {
        public int ComponenteID { get; set; }
        public string Funcionario { get; set; }              // opcional (podes enviar null)
        public int ClasseID { get; set; } = -1;
        public int? Fornecedor { get; set; }                 // opcional
        public int? SubEmpID { get; set; }                   // opcional
        public decimal NumHorasTrabalho { get; set; }        // obrigatório
        public decimal? NumHorasOrdem { get; set; }          // opcional
        public decimal? NumHorasAvariada { get; set; }       // opcional
        public decimal? PrecoUnit { get; set; }              // opcional (default 0)
        public Guid? ItemId { get; set; }                    // opcional (se tiveres ID externo)
    }

    public class ParteDiariaEquipamentoRequestDto
    {
        public ParteDiariaEquipamentoCabecalhoDto Cabecalho { get; set; }
        public List<ParteDiariaEquipamentoItemDto> Itens { get; set; }
    }


    public class FeriasFuncionarioModel
    {
        public string Funcionario { get; set; }
        public DateTime DataFeria { get; set; }
        public int EstadoGozo { get; set; } = 0;
        public int OriginouFalta { get; set; } = 1;
        public int TipoMarcacao { get; set; } = 1;
        public int OriginouFaltaSubAlim { get; set; } = 1;
        public decimal Duracao { get; set; } = 1.0M;
        public int Acerto { get; set; } = 0;
        public int? NumProc { get; set; }
        public int Origem { get; set; } = 0;
    }




}
