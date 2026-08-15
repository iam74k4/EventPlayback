from __future__ import annotations

import pytest

from eventplayback.ui import theme


def test_every_colour_token_is_a_light_dark_pair():
    tokens = [
        getattr(theme, name)
        for name in dir(theme)
        if name.isupper() and isinstance(getattr(theme, name), tuple)
    ]
    assert tokens, "expected the module to expose colour tokens"
    for token in tokens:
        assert len(token) == 2
        for half in token:
            assert half.startswith("#") and len(half) == 7


def test_blending_all_the_way_reaches_the_target():
    assert theme.blend(theme.TEXT, theme.WINDOW, 1.0) == theme.WINDOW


def test_blending_none_of_the_way_changes_nothing():
    assert theme.blend(theme.TEXT, theme.WINDOW, 0.0) == theme.TEXT


def test_blend_is_clamped_rather_than_extrapolated():
    assert theme.blend(theme.TEXT, theme.WINDOW, 5.0) == theme.WINDOW
    assert theme.blend(theme.TEXT, theme.WINDOW, -5.0) == theme.TEXT


@pytest.mark.parametrize("amount", [0.2, 0.55, 0.9])
def test_dimming_moves_towards_the_window_without_reaching_it(amount):
    dimmed = theme.dim(theme.STATUS_RECORDING, amount)
    assert dimmed != theme.STATUS_RECORDING
    assert dimmed != theme.WINDOW


def test_every_banner_severity_has_readable_text():
    assert set(theme.BANNER_TEXT) == {"info", "success", "warning", "error"}
