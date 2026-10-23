# syntax=docker/dockerfile:1

# Three stages, because what builds this is not what runs it: the compiler and
# its types stay behind, and the runtime carries the built entry and the
# production dependencies only.

# ---------------------------------------------------------------------------
# The build: TypeScript in, dist/ out.
# ---------------------------------------------------------------------------
FROM node:24-alpine AS build
WORKDIR /app

# The sources are copied before the install and the install ignores scripts:
# `prepare` runs `npm run build`, which would otherwise run against an empty
# `src`, and nothing in this dependency tree compiles anything natively, so a
# script that never runs is a supply-chain saving rather than a loss.
COPY package.json package-lock.json ./
RUN npm ci --ignore-scripts

COPY tsconfig.json tsconfig.build.json ./
COPY src ./src
RUN npm run build

# ---------------------------------------------------------------------------
# Dependencies as they run: no dev dependencies, no build tools.
# ---------------------------------------------------------------------------
FROM node:24-alpine AS deps
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --ignore-scripts

# ---------------------------------------------------------------------------
# Runtime.
# ---------------------------------------------------------------------------
FROM node:24-alpine AS runtime

# dumb-init is PID 1 for a reason. Node as PID 1 ignores SIGTERM, so `docker
# stop` would wait out the timeout and kill it; and the sign-in flow detaches a
# child that outlives its request, which somebody has to reap.
RUN apk add --no-cache dumb-init

ENV NODE_ENV=production \
    HOME=/data \
    # The session, the sign-in state and the downloaded attachments. The image
    # ships none of them; /data is the volume, so `docker run -v` is what
    # decides where they live on the host.
    OUTLOOK_MCP_HOME=/data \
    OUTLOOK_MCP_ATTACHMENT_DIR=/data/attachments \
    # A server bound to 127.0.0.1 *inside* the container cannot be reached
    # through a published port: the traffic arrives on the container's own
    # interface, not on its loopback. So the bind is a wildcard and the loopback
    # rule moves to the host side of `-p 127.0.0.1:3939:3939`, which is where it
    # keeps other machines out. Host and Origin are checked on every request
    # whatever this is set to, so a name that is not the machine's own is
    # refused from either side.
    OUTLOOK_MCP_HOST=0.0.0.0 \
    OUTLOOK_MCP_PORT=3939

WORKDIR /app
COPY --from=deps --chown=node:node /app/node_modules ./node_modules
COPY --from=build --chown=node:node /app/dist ./dist
COPY --chown=node:node package.json README.md SECURITY.md CHANGELOG.md LICENSE ./
COPY --chown=node:node skills ./skills
COPY --chown=node:node docs/clients.md ./docs/clients.md

# 0700 and owned by `node`: a mounted volume inherits this ownership the first
# time it is used, and the token cache inside it is written 0600.
RUN mkdir -p /data/attachments && chown -R node:node /data && chmod 700 /data /data/attachments

# `docker exec` starts a program, not the entrypoint, so `docker exec smoke
# outlook login` needs the entry on the PATH under a name a person will type.
RUN printf '#!/bin/sh\nexec node /app/dist/index.js "$@"\n' > /usr/local/bin/outlook \
  && chmod 755 /usr/local/bin/outlook

USER node

EXPOSE 3939
VOLUME ["/data"]

# The one route that answers without touching Microsoft. It is what `docker ps`
# reports for a `serve --http` container; a container running `login` has no
# listener and will read as unhealthy, which is only true of it.
HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
  CMD node -e "const p = process.env.OUTLOOK_MCP_PORT || 3939; fetch('http://127.0.0.1:' + p + '/health').then((r) => process.exit(r.ok ? 0 : 1), () => process.exit(1))"

# The entry is fixed, so anything after the image name is an outlook command:
# `docker run … login`, `docker exec outlook-mcp outlook emails list`.
ENTRYPOINT ["dumb-init", "--", "outlook"]
CMD ["serve", "--http"]
