import 'dart:async';

import 'package:flutter/material.dart';

import '../screens/login_screen.dart';
import '../services/api_service.dart';
import 'inscricoes_screen.dart';
import 'relatorios_screen.dart';
import 'lista_presenca_screen.dart';

class AdminHome extends StatefulWidget {
  const AdminHome({
    super.key,
  });

  @override
  State<AdminHome> createState() => _AdminHomeState();
}

class _AdminHomeState extends State<AdminHome>
    with WidgetsBindingObserver {
  Timer? _heartbeat;

  String _codigoAgenteTeste = '';
  bool _travaTesteAtiva = false;

  @override
  void initState() {
    super.initState();

    WidgetsBinding.instance.addObserver(this);

    _iniciarHeartbeat();
    _carregarStatusTrava();
  }

  @override
  void dispose() {
    _heartbeat?.cancel();

    WidgetsBinding.instance.removeObserver(this);

    super.dispose();
  }

  Future<void> _carregarStatusTrava() async {
    try {
      final resultado =
          await ApiService.obterStatusTravaTeste();

      if (!mounted) return;

      if (resultado['success'] != true) {
        if (resultado['sessaoExpirada'] == true) {
          _sessaoEncerrada();
        }

        return;
      }

      final ativa =
          resultado['ativa'] == true;

      final codigo =
          resultado['codigo']?.toString() ?? '';

      setState(() {
        _travaTesteAtiva = ativa;
        _codigoAgenteTeste = codigo;
      });
    } catch (e) {
      debugPrint(
        'Erro ao carregar trava de teste: $e',
      );
    }
  }

  void _iniciarHeartbeat() {
    _heartbeat?.cancel();

    _heartbeat = Timer.periodic(
      const Duration(seconds: 60),
      (_) async {
        if (!ApiService.possuiSessao) {
          _sessaoEncerrada();
          return;
        }

        final ok =
            await ApiService.heartbeat();

        if (!ok) {
          _sessaoEncerrada();
        }
      },
    );
  }

  @override
  void didChangeAppLifecycleState(
    AppLifecycleState state,
  ) {
    if (state ==
        AppLifecycleState.resumed) {
      if (ApiService.possuiSessao) {
        ApiService.heartbeat();

        _iniciarHeartbeat();

        _carregarStatusTrava();
      }
    } else if (state ==
            AppLifecycleState.paused ||
        state ==
            AppLifecycleState.inactive) {
      _heartbeat?.cancel();
    }
  }

  void _sessaoEncerrada() {
    if (!mounted) return;

    _heartbeat?.cancel();

    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(
        backgroundColor: Colors.red,
        content: Text(
          'Sua sessão foi encerrada. Faça login novamente.',
        ),
      ),
    );

    Navigator.of(context).pushAndRemoveUntil(
      MaterialPageRoute(
        builder: (_) =>
            const LoginScreen(),
      ),
      (route) => false,
    );
  }

  Widget botao({
    required IconData icone,
    required String titulo,
    required VoidCallback onTap,
  }) {
    return SizedBox(
      width: double.infinity,
      height: 65,
      child: ElevatedButton.icon(
        icon: Icon(icone),
        label: Text(
          titulo,
          style: const TextStyle(
            fontSize: 18,
            fontWeight: FontWeight.bold,
          ),
        ),
        onPressed: onTap,
      ),
    );
  }

  void _mostrarMensagem(
    String mensagem, {
    bool erro = false,
  }) {
    if (!mounted) return;

    ScaffoldMessenger.of(context)
        .showSnackBar(
      SnackBar(
        backgroundColor:
            erro ? Colors.red : null,
        content: Text(mensagem),
      ),
    );
  }

  Future<void> _bloquearAgente(
    String codigo,
  ) async {
    final codigoNormalizado =
        codigo.trim();

    if (codigoNormalizado.isEmpty) {
      _mostrarMensagem(
        'Informe o código do agente fictício.',
        erro: true,
      );

      return;
    }

    try {
      final resultado =
          await ApiService.bloquearAgenteTeste(
        codigoNormalizado,
      );

      if (!mounted) return;

      if (resultado['success'] != true) {
        if (resultado['sessaoExpirada'] ==
            true) {
          _sessaoEncerrada();
          return;
        }

        _mostrarMensagem(
          resultado['mensagem']?.toString() ??
              'Não foi possível ativar a trava.',
          erro: true,
        );

        return;
      }

      setState(() {
        _codigoAgenteTeste =
            codigoNormalizado;

        _travaTesteAtiva = true;
      });

      _mostrarMensagem(
        'Trava ativada para o agente $codigoNormalizado.',
      );
    } catch (e) {
      _mostrarMensagem(
        e.toString().replaceFirst(
              'Exception: ',
              '',
            ),
        erro: true,
      );
    }
  }

  Future<void> _restaurarAgente(
    String codigo,
  ) async {
    final codigoNormalizado =
        codigo.trim();

    if (codigoNormalizado.isEmpty) {
      _mostrarMensagem(
        'Informe o código do agente fictício.',
        erro: true,
      );

      return;
    }

    try {
      final resultado =
          await ApiService.restaurarAgenteTeste(
        codigoNormalizado,
      );

      if (!mounted) return;

      if (resultado['success'] != true) {
        if (resultado['sessaoExpirada'] ==
            true) {
          _sessaoEncerrada();
          return;
        }

        _mostrarMensagem(
          resultado['mensagem']?.toString() ??
              'Não foi possível restaurar o agente.',
          erro: true,
        );

        return;
      }

      setState(() {
        _codigoAgenteTeste =
            codigoNormalizado;

        _travaTesteAtiva = false;
      });

      _mostrarMensagem(
        'Agente $codigoNormalizado restaurado.',
      );
    } catch (e) {
      _mostrarMensagem(
        e.toString().replaceFirst(
              'Exception: ',
              '',
            ),
        erro: true,
      );
    }
  }

  Future<void> _abrirPainelTeste() async {
    await _carregarStatusTrava();

    if (!mounted) return;

    final controlador =
        TextEditingController(
      text: _codigoAgenteTeste,
    );

    await showDialog<void>(
      context: context,
      builder: (dialogContext) {
        return StatefulBuilder(
          builder: (
            context,
            setDialogState,
          ) {
            return AlertDialog(
              title: const Text(
                'TESTE DE CONEXÃO',
                style: TextStyle(
                  fontWeight: FontWeight.bold,
                ),
              ),
              content:
                  SingleChildScrollView(
                child: Column(
                  mainAxisSize:
                      MainAxisSize.min,
                  crossAxisAlignment:
                      CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'Agente fictício:',
                      style: TextStyle(
                        fontWeight:
                            FontWeight.bold,
                      ),
                    ),
                    const SizedBox(
                      height: 10,
                    ),
                    TextField(
                      controller:
                          controlador,
                      keyboardType:
                          TextInputType.number,
                      decoration:
                          const InputDecoration(
                        border:
                            OutlineInputBorder(),
                        labelText:
                            'Código do agente',
                        hintText:
                            'Ex.: 9999',
                      ),
                    ),
                    const SizedBox(
                      height: 25,
                    ),
                    Container(
                      width: double.infinity,
                      padding:
                          const EdgeInsets.all(
                        15,
                      ),
                      decoration:
                          BoxDecoration(
                        borderRadius:
                            BorderRadius.circular(
                          10,
                        ),
                        border: Border.all(
                          color: _travaTesteAtiva
                              ? Colors.red
                              : Colors.green,
                        ),
                      ),
                      child: Column(
                        crossAxisAlignment:
                            CrossAxisAlignment
                                .start,
                        children: [
                          const Text(
                            'STATUS',
                            style:
                                TextStyle(
                              fontWeight:
                                  FontWeight.bold,
                            ),
                          ),
                          const SizedBox(
                            height: 8,
                          ),
                          Text(
                            _travaTesteAtiva
                                ? 'BLOQUEADO'
                                : 'NORMAL',
                            style:
                                TextStyle(
                              fontSize: 18,
                              fontWeight:
                                  FontWeight
                                      .bold,
                              color:
                                  _travaTesteAtiva
                                      ? Colors.red
                                      : Colors.green,
                            ),
                          ),
                          const SizedBox(
                            height: 8,
                          ),
                          Text(
                            _travaTesteAtiva
                                ? 'As operações do agente fictício retornarão "Falha de conexão".'
                                : 'O agente fictício está funcionando normalmente.',
                            style:
                                const TextStyle(
                              color:
                                  Colors.black54,
                            ),
                          ),
                        ],
                      ),
                    ),
                    const SizedBox(
                      height: 20,
                    ),
                    SizedBox(
                      width:
                          double.infinity,
                      height: 55,
                      child:
                          ElevatedButton.icon(
                        icon: const Icon(
                          Icons
                              .signal_wifi_off,
                        ),
                        label: const Text(
                          'BLOQUEAR AGENTE',
                          style:
                              TextStyle(
                            fontWeight:
                                FontWeight.bold,
                          ),
                        ),
                        style:
                            ElevatedButton
                                .styleFrom(
                          backgroundColor:
                              Colors.red,
                          foregroundColor:
                              Colors.white,
                        ),
                        onPressed:
                            () async {
                          final codigo =
                              controlador
                                  .text
                                  .trim();

                          await _bloquearAgente(
                            codigo,
                          );

                          if (!mounted) {
                            return;
                          }

                          setDialogState(
                            () {},
                          );
                        },
                      ),
                    ),
                    const SizedBox(
                      height: 12,
                    ),
                    SizedBox(
                      width:
                          double.infinity,
                      height: 55,
                      child:
                          ElevatedButton.icon(
                        icon: const Icon(
                          Icons
                              .signal_wifi_4_bar,
                        ),
                        label: const Text(
                          'RESTAURAR AGENTE',
                          style:
                              TextStyle(
                            fontWeight:
                                FontWeight.bold,
                          ),
                        ),
                        style:
                            ElevatedButton
                                .styleFrom(
                          backgroundColor:
                              Colors.green,
                          foregroundColor:
                              Colors.white,
                        ),
                        onPressed:
                            () async {
                          final codigo =
                              controlador
                                  .text
                                  .trim();

                          await _restaurarAgente(
                            codigo,
                          );

                          if (!mounted) {
                            return;
                          }

                          setDialogState(
                            () {},
                          );
                        },
                      ),
                    ),
                  ],
                ),
              ),
              actions: [
                TextButton(
                  onPressed: () {
                    Navigator.pop(
                      dialogContext,
                    );
                  },
                  child:
                      const Text(
                    'FECHAR',
                  ),
                ),
              ],
            );
          },
        );
      },
    );

    controlador.dispose();
  }

  Future<void> _logout() async {
    final confirmar =
        await showDialog<bool>(
      context: context,
      builder:
          (dialogContext) {
        return AlertDialog(
          title: const Text(
            'Sair do sistema',
          ),
          content: const Text(
            'Deseja realmente sair do painel administrativo?',
          ),
          actions: [
            TextButton(
              onPressed: () {
                Navigator.pop(
                  dialogContext,
                  false,
                );
              },
              child: const Text(
                'CANCELAR',
              ),
            ),
            ElevatedButton(
              onPressed: () {
                Navigator.pop(
                  dialogContext,
                  true,
                );
              },
              child: const Text(
                'SAIR',
              ),
            ),
          ],
        );
      },
    );

    if (confirmar != true) {
      return;
    }

    _heartbeat?.cancel();

    await ApiService.logout();

    if (!mounted) return;

    Navigator.of(context)
        .pushAndRemoveUntil(
      MaterialPageRoute(
        builder: (_) =>
            const LoginScreen(),
      ),
      (route) => false,
    );
  }

  @override
  Widget build(
    BuildContext context,
  ) {
    final desktop =
        MediaQuery.of(context)
                .size
                .width >
            800;

    return Scaffold(
      backgroundColor:
          const Color(0xff021426),
      appBar: AppBar(
        backgroundColor:
            const Color(0xff00162f),
        centerTitle: true,
        title: const Text(
          'Administrador',
        ),
        actions: [
          IconButton(
            icon:
                const Icon(Icons.logout),
            tooltip: 'Sair',
            onPressed: _logout,
          ),
        ],
      ),
      body: Center(
        child:
            SingleChildScrollView(
          child: Padding(
            padding:
                const EdgeInsets.all(
              30,
            ),
            child: ConstrainedBox(
              constraints:
                  const BoxConstraints(
                maxWidth: 550,
              ),
              child: Column(
                children: [
                  const Icon(
                    Icons
                        .admin_panel_settings,
                    size: 110,
                    color: Colors.white,
                  ),
                  const SizedBox(
                    height: 25,
                  ),
                  Text(
                    'Painel Administrativo',
                    textAlign:
                        TextAlign.center,
                    style:
                        TextStyle(
                      color:
                          Colors.white,
                      fontSize:
                          desktop
                              ? 32
                              : 26,
                      fontWeight:
                          FontWeight.bold,
                    ),
                  ),
                  const SizedBox(
                    height: 10,
                  ),
                  const Text(
                    'Gerenciamento do sistema de plantões',
                    textAlign:
                        TextAlign.center,
                    style:
                        TextStyle(
                      color:
                          Colors.white70,
                      fontSize: 16,
                    ),
                  ),
                  const SizedBox(
                    height: 50,
                  ),
                  botao(
                    icone:
                        Icons.list_alt,
                    titulo:
                        'VISUALIZAR INSCRIÇÕES',
                    onTap: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (_) =>
                              const InscricoesScreen(),
                        ),
                      );
                    },
                  ),
                  const SizedBox(
                    height: 20,
                  ),
                  botao(
                    icone:
                        Icons.picture_as_pdf,
                    titulo:
                        'RELATÓRIO GERAL PDF',
                    onTap: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (_) =>
                              const RelatoriosScreen(),
                        ),
                      );
                    },
                  ),
                  const SizedBox(
                    height: 20,
                  ),
                  botao(
                    icone:
                        Icons.fact_check,
                    titulo:
                        'LISTA DE PRESENÇA',
                    onTap: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (_) =>
                              const ListaPresencaScreen(),
                        ),
                      );
                    },
                  ),
                  const SizedBox(
                    height: 20,
                  ),
                  botao(
                    icone:
                        Icons.bug_report,
                    titulo:
                        'TESTE DE CONEXÃO',
                    onTap:
                        _abrirPainelTeste,
                  ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}