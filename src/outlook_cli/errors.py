"""Exception hierarchy for outlook-cli. Each exception carries an exit_code."""


class UserError(Exception):
    """Base for all user-facing errors. exit_code defaults to 1."""

    exit_code: int = 1


class SessionExpired(UserError):
    """Refresh token expired or revoked. User must re-run 'outlook login'."""

    exit_code = 77


class NotFound(UserError):
    """ID or index does not resolve to a Graph resource."""

    exit_code = 64
