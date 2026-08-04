from outlook_cli.render.redact import redact_secrets


def test_redacts_authorization_bearer_header() -> None:
    text = "Authorization: Bearer eyJ0eXAiOiJKV1Qi.abc.def-ghi"
    out = redact_secrets(text)
    assert "eyJ0eXAi" not in out
    assert "Bearer ***REDACTED***" in out


def test_redacts_access_token_in_json() -> None:
    text = '"access_token": "eyJ0eXAiOiJKV1Qi.abc.def-ghi"'
    out = redact_secrets(text)
    assert "eyJ0eXAi" not in out
    assert '"access_token": "***REDACTED***"' in out


def test_redacts_refresh_token_in_json() -> None:
    text = '"refresh_token": "1.AXYA.something-long-here"'
    out = redact_secrets(text)
    assert "AXYA" not in out
    assert '"refresh_token": "***REDACTED***"' in out


def test_redacts_form_encoded_refresh_token() -> None:
    text = "refresh_token=1.AXYA.something-long-here&other=foo"
    out = redact_secrets(text)
    assert "AXYA" not in out
    assert "refresh_token=***REDACTED***" in out


def test_preserves_non_secret_text() -> None:
    text = "GET /me/messages HTTP/1.1\nHost: graph.microsoft.com"
    assert redact_secrets(text) == text


def test_no_real_looking_jwt_ever_survives() -> None:
    fake_jwt = "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiIxIn0.signature_x"
    text = f"some prefix {fake_jwt} some suffix"
    out = redact_secrets(text)
    assert fake_jwt not in out
