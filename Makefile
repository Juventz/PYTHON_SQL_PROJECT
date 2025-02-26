# Variables
include .env
export $(shell sed 's/=.*//' .env)
COMPOSE=docker compose
DOCKER=docker
SERVICE=postgres
IMAGE=postgres:latest
CONTAINER_NAME=$(shell $(DOCKER) ps --format "{{.Names}}" | grep $(SERVICE))

# Lancer le conteneur en arrière-plan
up:
	$(COMPOSE) up -d
	@echo "✅ Conteneur PostgreSQL démarré."

# Arrêter le conteneur
down:
	$(COMPOSE) down
	@echo "🛑 Conteneur arrêté."

# Supprimer le conteneur et les volumes
clean:
	$(COMPOSE) down -v
	@echo "🗑️  Conteneur et volumes supprimés."

# Supprimer l'image PostgreSQL
rmi:
	$(DOCKER) rmi -f $(IMAGE)
	@echo "🗑️  Image PostgreSQL supprimée."

# Voir les logs du conteneur
logs:
	$(DOCKER) logs -f $(CONTAINER_NAME)

# Se connecter à PostgreSQL avec psql
psql:
	$(DOCKER) exec -it $(CONTAINER_NAME) psql -U $(POSTGRES_USER) -d $(POSTGRES_DB)

# Construire et lancer tout
rebuild: clean up
	@echo "🔄 Conteneur reconstruit."

# Afficher les conteneurs en cours d'exécution
ps:
	$(DOCKER) ps

# Afficher les images Docker
images:
	$(DOCKER) images

# Afficher les volumes Docker
volumes:
	$(DOCKER) volume ls
