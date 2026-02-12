"""
Configuration loading for the LogTen-to-CAAI pipeline.

Uses Python's built-in configparser (no extra dependencies).
Supports config.ini file with CLI argument overrides.
"""

import configparser
import os


DEFAULT_CONFIG = {
    'pilot': {
        'name': '',
    },
    'import': {
        'input_file': '',
        'format': 'auto',
        'column_mapping': '',
    },
    'files': {
        'logten_export': './Export Flights (Tab).txt',
        'logbook_output': './Flight_Logbook.xlsx',
        'caai_template': './templates/tofes-shaot-blank.xlsx',
        'caai_output': './CAAI_Tofes_Shaot_Filled.xlsx',
        'custom_airports': '',
    },
}


class Config:
    """Pipeline configuration."""

    def __init__(self):
        self.pilot_name = ''
        self.input_file = ''
        self.input_format = 'auto'
        self.column_mapping = ''
        self.logten_export = ''
        self.logbook_output = ''
        self.caai_template = ''
        self.caai_output = ''
        self.custom_airports = ''

    @classmethod
    def from_file(cls, config_path):
        """Load configuration from an INI file.

        Args:
            config_path: Path to the config.ini file.

        Returns:
            Config instance.
        """
        config = cls()
        parser = configparser.ConfigParser()

        # Set defaults
        for section, values in DEFAULT_CONFIG.items():
            parser[section] = values

        # Read user config
        if os.path.exists(config_path):
            parser.read(config_path, encoding='utf-8')

        # Resolve paths relative to config file directory
        config_dir = os.path.dirname(os.path.abspath(config_path))

        config.pilot_name = parser.get('pilot', 'name', fallback='')

        # Import settings
        config.input_format = parser.get('import', 'format', fallback='auto')

        for attr, section, key in [
            ('input_file', 'import', 'input_file'),
            ('column_mapping', 'import', 'column_mapping'),
            ('logten_export', 'files', 'logten_export'),
            ('logbook_output', 'files', 'logbook_output'),
            ('caai_template', 'files', 'caai_template'),
            ('caai_output', 'files', 'caai_output'),
            ('custom_airports', 'files', 'custom_airports'),
        ]:
            val = parser.get(section, key, fallback='')
            if val and not os.path.isabs(val):
                val = os.path.join(config_dir, val)
            setattr(config, attr, val)

        return config

    def override(self, **kwargs):
        """Override config values from CLI arguments.

        Only overrides non-None values.
        """
        for key, value in kwargs.items():
            if value is not None and hasattr(self, key):
                setattr(self, key, value)

    def get_import_source(self):
        """Determine import source and format.

        Returns:
            Tuple of (file_path, format_str) or (None, None) if no import configured.
        """
        if self.input_file:
            return self.input_file, self.input_format
        elif self.logten_export:
            return self.logten_export, 'logten'
        return None, None

    def validate(self, step=None):
        """Validate that required files exist for the given step.

        Args:
            step: Pipeline step name, or None for full pipeline.

        Raises:
            FileNotFoundError: If a required file is missing.
        """
        if step in (None, 'import', 'logbook'):
            source, fmt = self.get_import_source()
            if source and not os.path.exists(source):
                raise FileNotFoundError(
                    f"Input file not found: {source}\n"
                    f"Check the file path in your config.ini."
                )
            elif not source:
                raise FileNotFoundError(
                    "No input file configured.\n"
                    "Set input_file in [import] section or logten_export in [files] section."
                )

        if step in ('distances', 'caai-columns', 'fill-form', 'analyze'):
            if not os.path.exists(self.logbook_output):
                raise FileNotFoundError(
                    f"Logbook file not found: {self.logbook_output}\n"
                    f"Run the 'import' step first to create it."
                )

        if step in (None, 'fill-form'):
            if not os.path.exists(self.caai_template):
                raise FileNotFoundError(
                    f"CAAI form template not found: {self.caai_template}\n"
                    f"Place the blank tofes-shaot Excel file in the templates/ directory."
                )

    def __repr__(self):
        return (
            f"Config(\n"
            f"  pilot_name='{self.pilot_name}',\n"
            f"  input_file='{self.input_file}',\n"
            f"  input_format='{self.input_format}',\n"
            f"  logten_export='{self.logten_export}',\n"
            f"  logbook_output='{self.logbook_output}',\n"
            f"  caai_template='{self.caai_template}',\n"
            f"  caai_output='{self.caai_output}',\n"
            f")"
        )
