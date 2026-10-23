import { configureLogging } from "../src/core/log.ts";

/**
 * Loaded before every test file through `node --test --import`, so the suite runs
 * with the logs off.
 *
 * `info` is this project's default level, which means a test of an MCP tool would
 * otherwise print a line for the call, another for the Graph request under it and
 * a third for the result — and the line that matters, the one about the assertion
 * that failed, would be somewhere in that. `test/log.test.ts` sets a level and a
 * sink of its own, because what it is testing is the logging.
 *
 * `force` is what makes this stick: `serve`, `startHttp` and the CLI each apply
 * their own configuration, and one of them runs inside nearly every test here.
 */
configureLogging({ level: "silent", force: true });
