package br.com.pakmatic.listapadraoifs.service;

import br.com.pakmatic.listapadraoifs.model.ListaEntry;
import br.com.pakmatic.listapadraoifs.repository.ListaRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ListaService {

    @Autowired
    private ListaRepository listaRepository;

    public ListaEntry salvar(ListaEntry entry) {
        return listaRepository.save(entry);
    }

    public List<ListaEntry> listarTodos() {
        return listaRepository.findAll();
    }
}
