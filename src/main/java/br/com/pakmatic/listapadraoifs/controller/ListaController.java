package br.com.pakmatic.listapadraoifs.controller;

import br.com.pakmatic.listapadraoifs.model.ListaEntry;
import br.com.pakmatic.listapadraoifs.repository.ListaRepository;
import br.com.pakmatic.listapadraoifs.service.ListaService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/listas")
public class ListaController {

    @Autowired
    private ListaService listaService;

    @Autowired
    private ListaRepository listaRepository;

    @PostMapping
    public ResponseEntity<ListaEntry> salvar(@RequestBody ListaEntry entry) {
        ListaEntry salvo = listaService.salvar(entry);
        return ResponseEntity.ok(salvo);
    }

    @GetMapping
    public ResponseEntity<List<ListaEntry>> listarTodos() {
        return ResponseEntity.ok(listaService.listarTodos());
    }

    @DeleteMapping("/resetar")
    public ResponseEntity<Void> resetarLista() {
        listaRepository.deleteAll();
        return ResponseEntity.ok().build();
    }
}
