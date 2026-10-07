"""Quote-aware accessor/function parameter counts."""

import re

from ..types.type_inference import runtime_signature_parameter_text, split_signature_top_level


def member_parameter_counts(signature: str | None) -> tuple[int, int]:
    inner = runtime_signature_parameter_text(signature) if signature is not None else None
    if not inner or not inner.strip():
        return 0, 0
    parts = [part.strip() for part in split_signature_top_level(inner)]
    return len(parts), sum(not re.match(r"^(Optional|ParamArray)\b", part, re.IGNORECASE) and not part.startswith("[") for part in parts)
