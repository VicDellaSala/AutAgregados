"""Bounded CSV/ZIP input streams; never extract ZIP member paths."""
from contextlib import contextmanager, ExitStack
from pathlib import Path
import zipfile

MAX_UNCOMPRESSED_BYTES = 4 * 1024**3


@contextmanager
def open_r34(source):
    with ExitStack() as stack:
        stream = source if hasattr(source, 'read') else stack.enter_context(open(source, 'rb'))
        stream.seek(0)
        name = str(getattr(source, 'name', source))
        zipped = Path(name).suffix.lower() == '.zip' or zipfile.is_zipfile(stream)
        stream.seek(0)
        if not zipped:
            yield stream
            return
        try:
            archive = stack.enter_context(zipfile.ZipFile(stream))
            candidates = [member for member in archive.infolist()
                          if not member.is_dir() and member.filename.lower().endswith('.csv')
                          and not member.filename.startswith('__MACOSX/')]
            if len(candidates) != 1:
                raise ValueError(f'El ZIP del R34 debe contener un único CSV; se encontraron {len(candidates)}. Prepare un ZIP con el R34 que desea procesar.')
            member = candidates[0]
            if member.flag_bits & 1:
                raise ValueError('El ZIP del R34 tiene contraseña. Suba una copia sin contraseña.')
            if member.file_size > MAX_UNCOMPRESSED_BYTES:
                raise ValueError('El CSV dentro del ZIP supera el límite de 4 GB descomprimidos.')
            # ZipExtFile decompresses only what the chunk reader requests.
            # Passing an open member (not the ZIP path) avoids pandas loading
            # the entire uncompressed archive into memory.
            with archive.open(member) as csv:
                yield csv
        except (zipfile.BadZipFile, EOFError) as error:
            raise ValueError('El ZIP del R34 está dañado o incompleto. Vuelva a comprimirlo y subirlo.') from error
        except NotImplementedError as error:
            raise ValueError('Compresión ZIP no compatible. Use un ZIP estándar con compresión Deflate.') from error
