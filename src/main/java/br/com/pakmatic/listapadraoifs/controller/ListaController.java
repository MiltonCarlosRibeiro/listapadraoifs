package br.com.pakmatic.listapadraoifs.controller;

import br.com.pakmatic.listapadraoifs.model.ListaEntry;
import br.com.pakmatic.listapadraoifs.service.ListaService;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/listas")
@CrossOrigin(origins = "*")
public class ListaController {

    private final ListaService service;

    public ListaController(ListaService service) {
        this.service = service;
    }

    @PostMapping
    public ListaEntry salvar(@RequestBody ListaEntry entry) {
        return service.salvar(entry);
    }

    @GetMapping
    public List<ListaEntry> listar() {
        return service.listarTudo();
    }

    @DeleteMapping("/{id}")
    public void deletar(@PathVariable Long id) {
        service.deletar(id);
    }
}
