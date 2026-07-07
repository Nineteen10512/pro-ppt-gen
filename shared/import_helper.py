from __future__ import annotations

import importlib
from typing import Iterable


def import_shared(module_name: str, attrs: Iterable[str] | None = None):
    """Import a module from the bundled shared package.

    The Pro PPT skill may run as ``pro_ppt_gen`` or from a broader ``skills``
    namespace. This helper keeps old absolute ``shared.*`` imports working.
    """
    candidates = [
        f"pro_ppt_gen.shared.{module_name}",
        f"shared.{module_name}",
        f"skills.shared.{module_name}",
    ]
    last_error: ImportError | None = None
    for dotted in candidates:
        try:
            module = importlib.import_module(dotted)
            if attrs is None:
                return module
            return tuple(getattr(module, attr) for attr in attrs)
        except ImportError as exc:
            last_error = exc
        except AttributeError:
            raise
    raise ImportError(f"Unable to import shared module '{module_name}'") from last_error
