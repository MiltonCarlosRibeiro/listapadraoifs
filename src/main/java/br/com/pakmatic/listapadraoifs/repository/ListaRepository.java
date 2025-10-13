package br.com.pakmatic.listapadraoifs.repository;

import br.com.pakmatic.listapadraoifs.model.ListaEntry;
import org.springframework.context.annotation.Profile;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Profile("db")
@Repository
public interface ListaRepository extends JpaRepository<ListaEntry, Long> {
}
