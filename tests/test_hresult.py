"""Tests for dn38_solver.com.hresult — COM error decoding.

The decoder operates on stringified pywin32 com_error tuples, so each
test passes the exact tuple shape that win32com.client.com_error would
produce. No COM dependency.
"""
from __future__ import annotations

from dn38_solver.com.hresult import decode_com_error, format_decoded


class TestKnownHResults:
    def test_generic_vba_exception_0x80048028(self):
        # The RP Puma 2026-05-13 incident shape: DISP_E_EXCEPTION wrapper
        # with secondary 0x80048028 (the generic VBA "Exception occurred").
        raw = (
            "(-2147352567, 'Exception occurred.', "
            "(0, 'Microsoft Excel', 'Exception occurred.', '', 0, "
            "-2146788248), None)"
        )
        d = decode_com_error(raw)
        assert d.hresult == -2147352567
        assert d.secondary == -2146788248
        assert d.auto_recoverable is True
        assert "VBA" in d.summary or "macro" in d.summary.lower()

    def test_e_unexpected_not_recoverable(self):
        raw = (
            "(-2147352567, 'Exception occurred.', "
            "(0, 'Microsoft Excel', '', '', 0, -2147418113), None)"
        )
        d = decode_com_error(raw)
        assert d.secondary == -2147418113
        assert d.auto_recoverable is False
        assert "E_UNEXPECTED" in d.summary or "unexpected" in d.summary.lower()

    def test_rpc_server_unavailable_not_recoverable(self):
        raw = (
            "(-2147023174, 'The remote procedure call failed.', "
            "(0, None, None, None, 0, 0), None)"
        )
        d = decode_com_error(raw)
        assert d.hresult == -2147023174
        assert d.auto_recoverable is False
        assert "RPC" in d.summary or "rpc" in d.summary.lower()

    def test_object_disconnected_not_recoverable(self):
        raw = (
            "(-2147417848, 'The object invoked has disconnected from its "
            "clients.', (0, None, None, None, 0, 0), None)"
        )
        d = decode_com_error(raw)
        assert d.hresult == -2147417848
        assert d.auto_recoverable is False
        assert "disconnected" in d.summary.lower()


class TestFallbacks:
    def test_unknown_secondary_under_dispatch_wrapper_is_recoverable(self):
        # DISP_E_EXCEPTION with a secondary we don't have a specific entry
        # for. Default policy: assume recoverable (re-import is cheap and
        # most generic-VBA errors clear on a clean re-import). The decoder
        # returns the secondary so the log line gives the operator a hint.
        raw = (
            "(-2147352567, 'Exception occurred.', "
            "(0, 'Microsoft Excel', '', '', 0, -2146826259), None)"
        )
        d = decode_com_error(raw)
        assert d.hresult == -2147352567
        assert d.secondary == -2146826259
        assert d.auto_recoverable is True

    def test_unknown_top_hresult_not_recoverable(self):
        raw = (
            "(-2147467259, 'Unspecified error', "
            "(0, None, None, None, 0, 0), None)"
        )
        d = decode_com_error(raw)
        assert d.hresult == -2147467259
        assert d.auto_recoverable is False

    def test_unparseable_raw_returns_fallback(self):
        d = decode_com_error("something not a com tuple at all")
        assert d.hresult is None
        assert d.auto_recoverable is False
        assert d.summary  # non-empty


class TestFormatting:
    def test_format_includes_hex_and_decimal(self):
        raw = (
            "(-2147352567, 'Exception occurred.', "
            "(0, 'Microsoft Excel', '', '', 0, -2146788248), None)"
        )
        d = decode_com_error(raw)
        text = format_decoded(d)
        # Both forms of the primary HRESULT should appear so operators can
        # match either grep style (some logs carry hex, some decimal).
        assert "0x80020009" in text  # -2147352567 in hex
        assert "-2147352567" in text
        # Secondary: -2146788248 = 0x800a9c68 (unsigned-32). Render checks
        # case-insensitive since format_decoded uses Python's lowercase hex.
        assert "0x800a9c68" in text.lower()
        assert "-2146788248" in text
        assert "Recovery" in text or "recovery" in text

    def test_format_recoverable_marker_present(self):
        raw = (
            "(-2147352567, 'Exception occurred.', "
            "(0, 'Microsoft Excel', '', '', 0, -2146788248), None)"
        )
        d = decode_com_error(raw)
        text = format_decoded(d)
        assert "Auto-recovery" in text

    def test_format_non_recoverable_omits_recovery_marker(self):
        raw = (
            "(-2147023174, 'The remote procedure call failed.', "
            "(0, None, None, None, 0, 0), None)"
        )
        d = decode_com_error(raw)
        text = format_decoded(d)
        assert "Auto-recovery: ELIGIBLE" not in text
