{
  description = "effect-spreadsheet: Effect v4 spreadsheet readers in pure TypeScript";

  inputs.nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";

  outputs = { self, nixpkgs }:
    let
      system = "x86_64-linux";
      pkgs = import nixpkgs { inherit system; };
    in {
      # Dev shell with the JS toolchain. Enter with: nix develop
      devShells.${system}.default = pkgs.mkShell {
        packages = with pkgs; [ nodejs_24 pnpm ];
        shellHook = ''
          echo "effect-spreadsheet dev shell: node $(node -v), pnpm $(pnpm -v)"
        '';
      };
    };
}
