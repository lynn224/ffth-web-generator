import importlib
u = importlib.import_module('ui_technical_forms')
print('collect_fase2_from_state available:', hasattr(u, 'collect_fase2_from_state'))
print('auto_save available:', hasattr(u, 'auto_save'))
print('load_autosave available:', hasattr(u, 'load_autosave'))
# Do not call auto_save unless you want a file created in history_database/autosave/
