package com.example.demo.entity;

import java.util.Set;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.ManyToOne;
import javax.persistence.OneToMany;

@Entity
public class Facture {
	@Id
	@GeneratedValue(strategy = GenerationType.AUTO)
	private Long id;
	
	@ManyToOne
	private Client client;
	
	@OneToMany(mappedBy = "facture")
	private Set<LigneFacture> ligneFactures;

	public Long getId() {
		return id;
	}

	public void setId(Long id) {
		this.id = id;
	}

	public Client getClient() {
		return client;
	}

	public void setClient(Client client) {
		this.client = client;
	}

	public Set<LigneFacture> getLigneFactures() {
		return ligneFactures;
	}

	public void setLigneFactures(Set<LigneFacture> ligneFacture) {
		this.ligneFactures = ligneFacture;
	}
	
	public Double getTotal() {
		Double total = 0.0;
		
		for(LigneFacture ligneFacture : ligneFactures) {
			total = total + ligneFacture.getSousTotal();
		}
		
		return total;
	}
	
}
