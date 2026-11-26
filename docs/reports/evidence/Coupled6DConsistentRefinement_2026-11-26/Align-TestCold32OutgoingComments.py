"""ДЛЯ ТЕСТОВ. Синхронизирует только расположение одинаковых комментариев, перенесённых VBE при добавлении нового тестового входа."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Align-TestCold32DirectionalComments.py'
code=file.read_text('utf-8-sig').replace('Cold899Cold32DirectionalJacobian_20261126','Cold900Cold32OutgoingJacobian_20261126')
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'cold32_outgoing_comments_alignment'})
