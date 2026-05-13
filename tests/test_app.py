import pytest
from doc_intelligence.web.app import (
    create_app,
    _arm_shutdown,
    _cancel_shutdown,
    _connected_clients,
    _doc_cache,
)


def test_create_app_returns_flask_app():
    app, socketio = create_app(testing=True, db_path=":memory:")
    assert app is not None
    assert app.testing is True
    assert socketio is not None


def test_create_app_registers_api_blueprint():
    app, _ = create_app(testing=True, db_path=":memory:")
    routes = {rule.rule for rule in app.url_map.iter_rules()}
    assert "/api/documents" in routes
    assert "/api/documents/scan" in routes
    assert "/api/documents/<doc_id>/parse" in routes


def test_doc_cache_module_level():
    _doc_cache.clear()
    assert isinstance(_doc_cache, dict)
    assert len(_doc_cache) == 0


def test_shutdown_arm_cancel_cycle():
    _connected_clients.clear()
    _cancel_shutdown()
    import doc_intelligence.web.app as app_module
    original_grace = app_module._SHUTDOWN_GRACE
    app_module._SHUTDOWN_GRACE = 5.0
    try:
        _arm_shutdown()
        assert app_module._shutdown_timer is not None
        _cancel_shutdown()
        assert app_module._shutdown_timer is None
    finally:
        app_module._SHUTDOWN_GRACE = original_grace


def test_shutdown_not_armed_when_clients_present():
    _connected_clients.clear()
    _cancel_shutdown()
    _connected_clients.add("sid-1")
    try:
        _arm_shutdown()
        import doc_intelligence.web.app as app_module
        assert app_module._shutdown_timer is None
    finally:
        _connected_clients.discard("sid-1")
        _cancel_shutdown()
