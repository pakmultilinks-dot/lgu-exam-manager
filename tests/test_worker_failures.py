import queue,time,logging
from types import SimpleNamespace
from app import App


def fake_app():
    return SimpleNamespace(job_busy=False,_jobs=queue.Queue(),title=lambda value:None,after=lambda *args:None,_poll_jobs=lambda:None)


def test_empty_exception_is_never_reported_as_success(monkeypatch):
    import app
    fake=fake_app();success=[];errors=[]
    monkeypatch.setattr(app.messagebox,'showerror',lambda title,text:errors.append(text))
    def fail():raise RuntimeError()
    assert App.run_job(fake,'Negative-control export',fail,lambda value:success.append(value))
    record=fake._jobs.get(timeout=5);fake._jobs.put(record)
    App._poll_jobs(fake)
    assert not fake.job_busy and not success
    assert len(errors)==1 and 'could not complete' in errors[0] and 'RuntimeError' not in errors[0] and 'Negative-control export' in errors[0]


def test_completion_callback_failure_is_not_logged_as_success(monkeypatch,caplog):
    fake=fake_app();errors=[]
    fake.report_callback_exception=lambda *args:errors.append(args)
    def callback(value):raise ValueError('Simulated completion failure')
    fake._jobs.put(('Test callback',callback,42,None,.01))
    with caplog.at_level(logging.INFO,logger='lgu'):App._poll_jobs(fake)
    assert errors and not fake.job_busy
    assert 'success=False' in caplog.text and 'success=True' not in caplog.text


def test_successful_worker_delivers_result_once():
    fake=fake_app();results=[]
    App.run_job(fake,'Test result',lambda:42,results.append)
    record=fake._jobs.get(timeout=5);fake._jobs.put(record)
    App._poll_jobs(fake)
    assert results==[42] and not fake.job_busy


def test_missing_pdf_reader_does_not_turn_generated_preview_into_app_failure(tmp_path,monkeypatch):
    import course_ui
    target=tmp_path/'preview.pdf';target.write_bytes(b'%PDF-1.4\n%%EOF')
    notices=[]
    def unavailable(path):raise OSError('No PDF association')
    monkeypatch.setattr(course_ui.os,'startfile',unavailable)
    monkeypatch.setattr(course_ui.messagebox,'showinfo',lambda title,text:notices.append(text))
    course_ui.open_pdf_preview(str(target))
    assert target.exists() and len(notices)==1 and str(target) in notices[0]
