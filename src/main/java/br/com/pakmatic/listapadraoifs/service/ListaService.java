package br.com.pakmatic.listapadraoifs.service;

import br.com.pakmatic.listapadraoifs.model.ListaEntry;
import br.com.pakmatic.listapadraoifs.repository.ListaRepository;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ListaService {

    private final ListaRepository repository;

    public ListaService(ListaRepository repository) {
        this.repository = repository;
    }

    public ListaEntry salvar(ListaEntry entry) {
        return repository.save(entry);
    }

    public List<ListaEntry> listarTudo() {
        return repository.findAll();
    }

    public void deletar(Long id) {
        repository.deleteById(id);
    }
}
