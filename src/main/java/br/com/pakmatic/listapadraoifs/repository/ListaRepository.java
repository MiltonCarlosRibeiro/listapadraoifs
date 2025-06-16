package br.com.pakmatic.listapadraoifs.repository;

import br.com.pakmatic.listapadraoifs.model.ListaEntry;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ListaRepository extends JpaRepository<ListaEntry, Long> {
}
