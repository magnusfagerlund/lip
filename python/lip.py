from lip.commands.install import install
import argparse
import datetime
import os
import sys

import logging


def set_up_logging():
    '''
    Sets upp logging to both file and console 
    '''
    # set up logging to file 
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
                        datefmt='%Y-%m-%d %H:%M',
                        filename='./lip.log',
                        filemode='w')
    # define a Handler which writes INFO messages or higher to the sys.stderr
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    # set a format which is simpler for console use
    formatter = logging.Formatter('%(levelname)-8s %(message)s')
    # tell the handler to use this format
    console.setFormatter(formatter)
    # add the handler to the root logger
    logging.getLogger('').addHandler(console)


class LIP():

    def __init__(self):
        parser = argparse.ArgumentParser(
            description="LIP is a package manager for LIME Pro",
            usage="""lip <command> [options]

Commands
  install                      Installs packages
  remove                       Removes a package
  freeze                       Outputs installed packages as 

            """
         
        )
        parser.add_argument('command', help='Subcommand to run')
        # parse_args defaults to [1:] for args, but we need to
        # exclude the rest of the args too, or validation will fail
        args = parser.parse_args(sys.argv[1:2])
        if not hasattr(self, args.command):
            print('Unrecognized command')
            parser.print_help()
            sys.exit(1)
        # use dispatch pattern to invoke method with same name
        getattr(self, args.command)()

    def install(self):
        parser = argparse.ArgumentParser(description='Installs new packages')
        parser.add_argument('package', help='Name of package to install')
        args = parser.parse_args(sys.argv[2:])
        install(args.package)
        


if __name__ == '__main__':
    set_up_logging()
    LIP()
