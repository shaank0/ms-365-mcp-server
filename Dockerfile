# Builds the gateway fork of ms-365-mcp-server from source and runs it in
# Streamable HTTP mode for the MCP Gateway's `ms365` delegated-Entra backend.
#
# This container holds NO credentials. In HTTP mode the server requires an
# `Authorization: Bearer <Graph access token>` header on every MCP request;
# the MCP Gateway (src/auth/graph-tokens.ts) mints and injects a per-user
# token on each call, so one instance safely serves every gateway user under
# their own M365 identity. Nothing is configured in the image beyond the
# server itself.
FROM node:24-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run generate
RUN npm run build

FROM node:24-alpine AS release

WORKDIR /app

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/package*.json ./

ENV NODE_ENV=production
RUN npm ci --ignore-scripts --omit=dev

# Run as a non-root system user. `adduser -S` gives it a proper, writable
# home directory (/home/ms365), which the server needs at startup to create
# its log directory (see src/logger.ts) even though no per-user state is
# otherwise persisted here.
RUN addgroup -S ms365 && adduser -S ms365 -G ms365 && chown -R ms365:ms365 /app
USER ms365

EXPOSE 3000

ENTRYPOINT ["node", "dist/index.js"]

# Streamable HTTP transport on 3000, org-mode for Teams/SharePoint/shared
# mailboxes. These are CMD (default args), not baked into ENTRYPOINT, so
# `docker run <image> <other args>` can still override them if ever needed.
CMD ["--http", "3000", "--org-mode"]
