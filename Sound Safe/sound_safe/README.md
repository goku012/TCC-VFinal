# Monitor de Exposição Sonora (TCC) — modular por funções

Agora separado **por classes e por funções**, para reduzir arquivos gigantes:
- `app.py`: orquestração (fino).
- `ui_left.py` e `ui_right.py`: construção de UI.
- `settings_dialog.py`: modal de configurações.
- `monitor.py`: loop do monitor em funções puras.
- `audio.py`: backend PyCAW + enforcer de bloqueio.
- `charting.py`: desenho do gráfico.
- `reporting.py`: criação do Excel e estatísticas.
- `persistence.py`: leitura/gravação de settings.
- `gauge.py`, `utils.py`, `constants.py`, `com_guard.py`.

## Rodar
```bash
pip install -r requirements.txt
python -m sound_monitor.main
```

### Opcionais
- Excel: `pip install openpyxl`
- Windows volume control: `pip install pycaw comtypes`
