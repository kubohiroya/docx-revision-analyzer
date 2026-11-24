#!/usr/bin/env node
import { Command } from "commander";
import "../node/locale";
import { initLangFromArgv, t } from "../lib/i18n";
import { langOption } from "./common";
import { configureChartCommand } from "./chart";
import { configureFlowCommand } from "./flow";
import { configureScoreCommand } from "./score";
import { configureSnapshotCommand } from "./snapshot";
import { configureVersionsCommand } from "./versions";

const argv = process.argv.slice(2);
initLangFromArgv(argv);

const program = new Command()
  .name("docx-revision-analyzer")
  .description(t("cliDescription"))
  .version(require("../../package.json").version)
  .addOption(langOption())
  .enablePositionalOptions()
  .addHelpCommand(false);

const commands = [
  ["chart", "cliChart", configureChartCommand],
  ["flow", "cliFlow", configureFlowCommand],
  ["score", "cliScore", configureScoreCommand],
  ["snapshot", "cliSnapshot", configureSnapshotCommand],
  ["versions", "cliVersions", configureVersionsCommand],
] as const;

// Read settings only for the selected command, including when requesting help.
const initialized = new Set<Command>();
function initialize(command: Command): void {
  const entry = commands.find(([name]) => name === command.name());
  if (entry && !initialized.has(command)) {
    entry[2](command, argv);
    initialized.add(command);
  }
}

for (const [name, description] of commands) {
  program.command(name).description(t(description));
}
program.hook("preSubcommand", (_parent, command) => initialize(command));
program.command("help [command]")
  .description(t("cliHelp"))
  .addOption(langOption())
  .action((name?: string) => {
    if (!name) {
      program.outputHelp();
      return;
    }
    const command = program.commands.find((c) => c.name() === name);
    if (!command) {
      program.error(t("cliUnknownCommand", name));
      return;
    }
    initialize(command);
    command.outputHelp();
  });

program.parseAsync(argv, { from: "user" }).catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
