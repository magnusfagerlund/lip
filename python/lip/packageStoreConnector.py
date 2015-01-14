from enum import Enum
import requests
from .package import Package

class PackageStoreConnector():
    """
    A Package Store is a location where you can find LIME Pro packages.
    This class handles connections to a package store

    Attributes
    ----------
    @attribute store_path : Path to the store. Can be a URL or a file path.
    @attribute store_type : Type of the store. Can either be a web store or a file store

    """
    
    def __init__(self, store_path: str):
        """
        Search the store for a package and returns it

        Parameters
        ----------
        @param store_path: Path to the store. Can be a URL or a file path.

        """
        self.store_path = store_path
        self.store_type = self._parse_store_url()
    
    def find_package(self, package_name: str) -> Package:
        """
        Search the store for a package and returns it

        Parameters
        ----------
        @param package_name: name of package to search for

        Returns
        -------
        @return: A Package class object containing all package data
        """

        if self.store_type == self.StoreType.web:
            return self._fetch_from_web_store(package_name)
        else:
            return self._fetch_from_file_store(package_name)

    def _fetch_from_web_store(self, package_name: str) -> Package:
        """
        Helper function to search for a package on a web store

        Parameters
        ----------
        @param package_name: name of package to search for

        Returns
        -------
        @return: A Package class object containing all package data
        """
        r = requests.get(self.store_path + package_name + "/")

        if r.ok:
            return Package(r.json())
        else:
            return None

    def _fetch_from_file_store(self, package_name: str) -> Package:
        """
        Helper function to search for a package on a file store

        Parameters
        ----------
        @param package_name: name of package to search for

        Returns
        -------
        @return: A Package class object containing all package data
        """

        return None

    def _parse_store_url(self) -> self.StoreType:
        """
        Helper function to parse to supplied path to the store to decide which type of store we are dealing with

        Returns
        -------
        @return: A Enum to keep track of which store we are dealing with
        """
        if self.store_path[:4] == "http":
            return self.StoreType.web
        else:
            return self.StoreType.file

    class StoreType(Enum):
        web = 1
        file = 2