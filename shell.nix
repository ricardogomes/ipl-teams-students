{ pkgs ? import <nixpkgs> {} }:

pkgs.stdenv.mkDerivation {
  name = "powershell-env";
  buildInputs = [ pkgs.powershell ];
  shellHook = ''
    echo "PowerShell environment is ready."
  '';
}
