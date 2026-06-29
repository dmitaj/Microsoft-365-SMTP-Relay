#!/bin/bash
#
# deploy.sh - update the local clone from GitHub, build the Docker image,
# push it to the registry, and recreate the running container with the new
# image. Designed to run ON the Unraid/Docker host where the container lives.
#
# The container is managed by an Unraid template, so there is no compose file
# to "up". Instead we snapshot the existing container's exact `docker run`
# command (via the `runlike` helper), rebuild the image, then stop/remove and
# recreate the container from that snapshot. All env vars, volumes, ports, and
# the restart policy are preserved because they come straight from the running
# container.
#
# Usage:
#   ./deploy.sh                 # pull, build, push, recreate
#   ./deploy.sh --no-push       # skip the registry push (local build only)
#   ./deploy.sh --no-restart    # build/push but leave the container alone
#
# Override any of these with environment variables if your names differ:
#   IMAGE, CONTAINER, BRANCH, REPO_DIR

set -euo pipefail

# --- configuration -----------------------------------------------------------
IMAGE="${IMAGE:-ghcr.io/dmitaj/microsoft-365-smtp-relay}"
CONTAINER="${CONTAINER:-Microsoft-365-SMTP-Relay}"
BRANCH="${BRANCH:-main}"
# Default to the directory this script lives in (your clone).
REPO_DIR="${REPO_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)}"

DO_PUSH=1
DO_RESTART=1
for arg in "$@"; do
  case "$arg" in
    --no-push)    DO_PUSH=0 ;;
    --no-restart) DO_RESTART=0 ;;
    *) echo "Unknown option: $arg" >&2; exit 2 ;;
  esac
done

log() { echo -e "\n\033[1;34m==>\033[0m $*"; }

# Retry a command up to 4 times with exponential backoff (for flaky network).
retry() {
  local n=1 max=4 delay=2
  until "$@"; do
    if (( n >= max )); then
      echo "Command failed after $n attempts: $*" >&2
      return 1
    fi
    echo "Attempt $n failed; retrying in ${delay}s..." >&2
    sleep "$delay"; delay=$(( delay * 2 )); n=$(( n + 1 ))
  done
}

# --- 1. update the local clone ----------------------------------------------
log "Updating clone in $REPO_DIR ($BRANCH)"
cd "$REPO_DIR"
retry git fetch origin "$BRANCH"
git checkout "$BRANCH"
git reset --hard "origin/$BRANCH"

TAG="$(git rev-parse --short HEAD)"
log "Building $IMAGE:latest (and :$TAG) from $(git log -1 --oneline)"

# --- 2. build ----------------------------------------------------------------
docker build -t "$IMAGE:latest" -t "$IMAGE:$TAG" "$REPO_DIR"

# --- 3. push -----------------------------------------------------------------
if (( DO_PUSH )); then
  log "Pushing to registry"
  if ! retry docker push "$IMAGE:latest"; then
    echo "Push failed. Are you logged in?  echo \$TOKEN | docker login ghcr.io -u <user> --password-stdin" >&2
    exit 1
  fi
  retry docker push "$IMAGE:$TAG"
else
  log "Skipping push (--no-push)"
fi

# --- 4. recreate the container ----------------------------------------------
if (( ! DO_RESTART )); then
  log "Skipping container recreate (--no-restart). Done."
  exit 0
fi

if ! docker inspect "$CONTAINER" >/dev/null 2>&1; then
  echo "Container '$CONTAINER' not found - start it once from the Unraid template, then re-run." >&2
  exit 1
fi

log "Snapshotting current run command for '$CONTAINER'"
# runlike reconstructs the exact `docker run ...` for the existing container,
# preserving every env var, volume, port, label and the restart policy.
RUN_CMD="$(docker run --rm -v /var/run/docker.sock:/var/run/docker.sock \
  assaflavie/runlike "$CONTAINER")"

log "Stopping and removing old container"
docker stop "$CONTAINER" >/dev/null
docker rm "$CONTAINER"  >/dev/null

log "Recreating '$CONTAINER' on the new image"
# Same image:tag (:latest) => uses the freshly built local image.
eval "$RUN_CMD"

# --- 5. report ---------------------------------------------------------------
sleep 2
STATUS="$(docker inspect -f '{{.State.Status}} (health: {{if .State.Health}}{{.State.Health.Status}}{{else}}n/a{{end}})' "$CONTAINER" 2>/dev/null || echo unknown)"
log "Done. '$CONTAINER' is now: $STATUS"
echo "    Image:  $IMAGE:$TAG"
echo "    Logs:   docker logs -f $CONTAINER"

# Optional: reclaim space from the old image layers.
# docker image prune -f >/dev/null
