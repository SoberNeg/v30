# v20.27 STARTUP WINDOW FIX

База: LED_Screen_Designer_v20_27_RZD_ASLS_AUTOFILL_DEBUG_FIX(1).py

Исправлено:
- стартовый выбор РЖД/МЦД запускается через `after()` после старта mainloop;
- `grab_set()` выполняется только после отображения Toplevel;
- добавлен резервный текстовый диалог РЖД/МЦД;
- ошибка стартового диалога больше не скрывается молча;
- при ошибке создаётся `startup_dialog_error.log`;
- сохранён Windows Tcl/Tk packaging fix.
