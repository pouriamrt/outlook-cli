import pytest

from outlook_cli.errors import NotFound, SessionExpired, UserError


def test_session_expired_carries_exit_code_77() -> None:
    exc = SessionExpired("Session expired. Run 'outlook login'.")
    assert exc.exit_code == 77


def test_not_found_carries_exit_code_64() -> None:
    exc = NotFound("Item AAA not found.")
    assert exc.exit_code == 64


def test_user_error_carries_exit_code_1() -> None:
    exc = UserError("Bad input.")
    assert exc.exit_code == 1


def test_exceptions_inherit_from_base() -> None:
    with pytest.raises(UserError):
        raise SessionExpired("x")
