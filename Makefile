CC = gcc
CFLAGS = $(shell pkg-config --cflags gtk+-3.0 libxml-2.0) -Isrc
LDFLAGS = $(shell pkg-config --libs gtk+-3.0) -lzip -lxml2

SRC = src/main.c src/data/data.c src/pptx/pptx_gen.c
OBJ = $(SRC:.c=.o)
TARGET = LivretCompetences

all: $(TARGET)

$(TARGET): $(OBJ)
	$(CC) -o $@ $(OBJ) $(LDFLAGS)

%.o: %.c
	$(CC) -c -o $@ $< $(CFLAGS)

clean:
	rm -f $(OBJ) $(TARGET)
