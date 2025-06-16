package br.com.pakmatic.listapadraoifs.model;

import jakarta.persistence.*;
import lombok.*;

@Entity
@Table(name = "listas")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ListaEntry {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String codigoMaterial;
    private String nivel;
    private String tipoEstrutura;
    private String linha;
    private String itemComponente;
    private String qtdeMontagem;
    private String unidadeMedida;
    private String fatorSucata;
}
