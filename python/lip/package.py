import json
import semantic_version
from enum import Enum


class Package():
    
    def __init__(self, package_json_data: json):
        self.data = package_json_data
        self.name = ""
        self.author = ""
        self.shortDesc = ""
        self.status = self.PackageStatus.beta
        self.versions = []
        self.dependencies = []
        self.install = self.Install

        
    def to_JSON(self) -> str:
    	return "hepp"

    class PackageStatus(Enum):
        release = 1
        beta = 2
        development = 3

    class Version():

        def __init__(self, version_json_data: json):
            self.version = semantic_version.Version("1.0.0")
            self.date = ""
            self.comment = ""

        def to_JSON(self):
            return ""

    class Dependency():

        def __init__(self, dependency_json_data: json):
            self.name = ""
            self.specification = semantic_version.Spec(">1.0.0")

        def to_JSON(self):
            return ""

    class Install():

        def __init__(self):
            self.localize = []
            self.vba = []
            self.sql = []
            self.tables = []

        class Localize():

            def __init__(self):
                self.owner = ""
                self.context = ""
                self.localname = []

        class LocalString():
            def __init__(self):
                self.language = ""
                self.text = ""

        class VBA():
            def __init__(self):
                self.name = ""
                self.relPath = ""

        class SQL():
            def __init__(self):
                self.name = ""
                self.relPath = ""

        class Table():
            def __init__(self):
                self.name = ""
                self.localname = []
                self.fields = []

            class Field():
                def __init__(self):
                    self.name = ""
                    self.type = self.FieldType.text
                    self.length = 234
                    self.localname = []

                    class FieldType(Enum):
                        text = 1
                        option = 2
                        set = 3
                        sql = 4
