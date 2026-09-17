set dotenv-load
set positional-arguments

@default:
  just --list --unsorted

mod task 'dev/tasks/fix-table-corruption/justfile'
