#ifndef DATA_H
#define DATA_H

typedef struct {
    char *domain;
    char *subdomain;
    char *text;
} Competence;

typedef struct {
    char *domain;
    char *subdomain;
    char *description;
} Domaine;

typedef struct {
    char *domain;
    char *color;
} DomainColor;

typedef struct {
    Competence *items;
    int count;
} CompetenceList;

typedef struct {
    Domaine *items;
    int count;
} DomaineList;

typedef struct {
    DomainColor *items;
    int count;
} DomainColorList;

typedef struct {
    char *domain;
    char *subdomain;
    char *text;
    char *timestamp;
} SelectedCompetence;

typedef struct {
    char *name;
    char *first_name;
    char *birth_date;
    char *photo_path;
} StudentInfo;

typedef struct {
    int completed;
    char *year;
    char *school;
    char *teachers;
    char *photo_path;
    char *bilan1;
    char *bilan2;
} SectionData;

CompetenceList load_competences(const char *filename);
DomaineList load_domaines(const char *filename);
DomainColorList load_couleurs(const char *filename);

void free_competences(CompetenceList list);
void free_domaines(DomaineList list);
void free_couleurs(DomainColorList list);

#endif
