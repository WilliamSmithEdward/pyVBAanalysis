"""Enum storage and assignment array identities from XLIDE 11.1.0."""

import re
from collections.abc import Callable, Set

from ..completion.member_access import MemberCompletionContext, resolve_known_object_assignment_type
from ..host.host_model import host_display_name, resolve_host_enum
from ..runtime.vba_runtime import resolve_vba_library_qualifier
from ..types.type_names import is_known_scalar_type, normalize_type


def create_assignment_coercion_type(ctx: MemberCompletionContext, source_enums: Set[str] = frozenset()) -> Callable[[str], str]:
    cache: dict[str, str] = {}
    project: dict[str, tuple[bool, bool]] | None = None

    def coercion(declared: str) -> str:
        nonlocal project
        if is_known_scalar_type(normalize_type(declared) or ""):
            return declared
        key = declared.strip().lower()
        if key in cache:
            return cache[key]
        if project is None:
            project = {}
            for surface in ctx.project_class_members or ():
                for name in (surface.name, f"{surface.module_name}.{surface.name}"):
                    enum, non_enum = project.get(name.lower(), (False, False))
                    project[name.lower()] = enum or surface.kind == "enum", non_enum or surface.kind != "enum"
        enum, non_enum = project.get(key, (False, False))
        result = declared
        if key in source_enums or (enum and not non_enum):
            result = "Long"
        elif not non_enum:
            runtime = resolve_vba_library_qualifier(re.sub(r"^VBA\.", "", declared, flags=re.IGNORECASE))
            enumeration = resolve_host_enum(declared.split(".")[-1], ctx.model)
            prefix = declared.rsplit(".", 1)[0].lower() if "." in declared else None
            if (runtime and any(constant.get("type") == runtime.name for constant in runtime.constants or ())) or (enumeration and (prefix is None or prefix == enumeration.get("library", host_display_name(ctx.model)).lower())):
                result = "Long"
        cache[key] = result
        return result

    return coercion


def array_element_identity(type_: str | None, ctx: MemberCompletionContext, coercion: Callable[[str], str]) -> str:
    value_type = coercion(re.sub(r"\s*\(\s*\)\s*$", "", type_ or "Variant"))
    object_type = resolve_known_object_assignment_type(value_type, ctx)
    scalar = normalize_type(value_type) or "variant"
    return f"{object_type.kind}:{object_type.key}" if object_type else "longlong" if scalar == "longptr" else scalar


def array_by_ref_identity(type_: str | None, ctx: MemberCompletionContext, source_enums: Set[str]) -> str:
    element = re.sub(r"\s*\(\s*\)\s*$", "", type_ or "Variant").strip()
    key = element.lower()
    if key in source_enums:
        return f"source-enum:{key.split('.')[-1]}"
    enumeration = resolve_host_enum(element.split(".")[-1], ctx.model)
    prefix = element.rsplit(".", 1)[0].lower() if "." in element else None
    if enumeration and (prefix is None or prefix == enumeration.get("library", host_display_name(ctx.model)).lower()):
        return f"host-enum:{enumeration.get('library', host_display_name(ctx.model))}.{enumeration.get('displayName')}".lower()
    runtime = resolve_vba_library_qualifier(re.sub(r"^VBA\.", "", element, flags=re.IGNORECASE))
    if runtime and any(constant.get("type") == runtime.name for constant in runtime.constants or ()):
        return f"runtime-enum:{runtime.name.lower()}"
    object_type = resolve_known_object_assignment_type(element, ctx)
    return f"{object_type.kind}:{object_type.key}" if object_type else normalize_type(element) or "variant"
