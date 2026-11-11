# -*- coding: utf-8 -*-
"""sak-fs：知识库文件读写 MCP 服务器（纯标准库，不依赖 Node）。

替代原先 `npx @modelcontextprotocol/server-filesystem`，避免装 Node / 首次联网下载。
所有路径都被限制在允许的根目录内，越界直接拒绝。

环境变量：
    SAK_FS_ROOTS   允许访问的根目录，多个用 os.pathsep 分隔；
                   未设置时用 SAK_KB_ROOT；再没有就用仓库下的 kb_template。
"""
import fnmatch
import os
import sys
from pathlib import Path

from fastmcp import FastMCP

mcp = FastMCP("sak-fs")


def _roots() -> list[Path]:
    raw = os.environ.get("SAK_FS_ROOTS") or os.environ.get("SAK_KB_ROOT") or ""
    items = [p for p in raw.split(os.pathsep) if p.strip()]
    if not items:
        items = [str(Path(__file__).resolve().parents[2] / "kb_template")]
    return [Path(p).expanduser().resolve() for p in items]


def _resolve(path: str) -> Path:
    p = Path(path).expanduser()
    if not p.is_absolute():
        p = _roots()[0] / p
    p = p.resolve()
    for root in _roots():
        if p == root or p.is_relative_to(root):
            return p
    raise PermissionError(f"路径不在允许范围内：{p}\n允许的根目录：{[str(r) for r in _roots()]}")


@mcp.tool()
def read_file(path: str, offset: int = 0, limit: int = 2000) -> dict:
    """读取文本文件。path 可为绝对路径或相对第一个根目录的路径。
    offset 为起始行号（0 起），limit 为最多返回行数，超出时 truncated=True。"""
    p = _resolve(path)
    if not p.is_file():
        return {"status": "error", "message": f"文件不存在：{p}"}
    lines = p.read_text(encoding="utf-8", errors="replace").splitlines()
    chunk = lines[offset:offset + limit]
    return {
        "status": "ok",
        "path": str(p),
        "total_lines": len(lines),
        "offset": offset,
        "truncated": offset + limit < len(lines),
        "content": "\n".join(chunk),
    }


@mcp.tool()
def write_file(path: str, content: str) -> dict:
    """写入文本文件（覆盖）。父目录不存在会自动创建，编码 UTF-8。"""
    p = _resolve(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(content, encoding="utf-8")
    return {"status": "ok", "path": str(p), "chars": len(content)}


@mcp.tool()
def list_directory(path: str = ".") -> dict:
    """列目录，返回名称、类型（dir/file）、字节数。"""
    p = _resolve(path)
    if not p.is_dir():
        return {"status": "error", "message": f"目录不存在：{p}"}
    entries = []
    for child in sorted(p.iterdir(), key=lambda x: (not x.is_dir(), x.name)):
        entries.append({
            "name": child.name,
            "type": "dir" if child.is_dir() else "file",
            "size": child.stat().st_size if child.is_file() else 0,
        })
    return {"status": "ok", "path": str(p), "entries": entries}


@mcp.tool()
def directory_tree(path: str = ".", depth: int = 2) -> dict:
    """按层级列出目录树，depth 为最大深度（默认 2，最多 5）。"""
    p = _resolve(path)
    depth = max(0, min(int(depth), 5))
    lines: list[str] = []

    def walk(cur: Path, level: int) -> None:
        if level > depth:
            return
        try:
            children = sorted(cur.iterdir(), key=lambda x: (not x.is_dir(), x.name))
        except PermissionError:
            return
        for child in children:
            lines.append("  " * level + child.name + ("/" if child.is_dir() else ""))
            if child.is_dir():
                walk(child, level + 1)

    walk(p, 0)
    return {"status": "ok", "path": str(p), "tree": "\n".join(lines), "entries": len(lines)}


@mcp.tool()
def search_files(pattern: str, path: str = ".") -> dict:
    """在目录下按文件名通配符递归查找，如 pattern="*制度*.md"，最多返回 200 条。"""
    p = _resolve(path)
    hits = []
    for cur, dirs, files in os.walk(p):
        dirs[:] = [d for d in dirs if d not in {".git", "__pycache__", "node_modules"}]
        for name in files:
            if fnmatch.fnmatch(name, pattern):
                full = Path(cur) / name
                hits.append({"path": str(full), "size": full.stat().st_size})
                if len(hits) >= 200:
                    return {"status": "ok", "count": len(hits), "truncated": True, "matches": hits}
    return {"status": "ok", "count": len(hits), "truncated": False, "matches": hits}


@mcp.tool()
def move_file(source: str, target: str) -> dict:
    """移动或重命名文件，目标父目录不存在会自动创建。"""
    src, dst = _resolve(source), _resolve(target)
    if not src.is_file():
        return {"status": "error", "message": f"源文件不存在：{src}"}
    dst.parent.mkdir(parents=True, exist_ok=True)
    src.replace(dst)
    return {"status": "ok", "source": str(src), "target": str(dst)}


@mcp.tool()
def get_file_info(path: str) -> dict:
    """返回文件/目录的类型、大小、修改时间。"""
    p = _resolve(path)
    if not p.exists():
        return {"status": "error", "message": f"路径不存在：{p}"}
    st = p.stat()
    return {
        "status": "ok",
        "path": str(p),
        "type": "dir" if p.is_dir() else "file",
        "size": st.st_size,
        "mtime": st.st_mtime,
    }


if __name__ == "__main__":
    mcp.run()
