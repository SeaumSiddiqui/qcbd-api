package org.project.app.program.orphan.service;

import org.modelmapper.ModelMapper;
import org.project.app.config.ModelMapperCustomizer;
import org.project.app.program.orphan.domain.*;
import org.project.app.program.orphan.dto.*;
import org.springframework.context.annotation.Configuration;

@Configuration
public class OrphanModelMapperCustomizer implements ModelMapperCustomizer {
    @Override
    public void customize(ModelMapper mapper) {

        //For create
        mapper.createTypeMap(OrphanApplicationCreateDTO.class, OrphanApplication.class)
                .addMappings(m -> {
                    m.skip(OrphanApplication::setId);
                    m.skip(OrphanApplication::setVersion);
                    m.skip(OrphanApplication::setCreatedBy);
                    m.skip(OrphanApplication::setLastReviewedBy);
                    m.skip(OrphanApplication::setCreatedAt);
                    m.skip(OrphanApplication::setLastModifiedAt);
                });

        // For update
        mapper.createTypeMap(OrphanApplicationUpdateDTO.class, OrphanApplication.class)
                .addMappings(m -> {
                    m.skip(OrphanApplication::setFamilyMembers);
                    m.skip(OrphanApplication::setVersion);
                    m.skip(OrphanApplication::setCreatedBy);
                    m.skip(OrphanApplication::setLastReviewedBy);
                    m.skip(OrphanApplication::setCreatedAt);
                    m.skip(OrphanApplication::setLastModifiedAt);
                });

        // Shared nested mappings
        mapper.createTypeMap(PrimaryInformationDTO.class, PrimaryInformation.class)
                .addMappings(m -> m.skip(PrimaryInformation::setId));

        mapper.createTypeMap(AddressDTO.class, Address.class)
                .addMappings(m -> m.skip(Address::setId));

        mapper.createTypeMap(BasicInformationDTO.class, BasicInformation.class)
                .addMappings(m -> m.skip(BasicInformation::setId));

        mapper.createTypeMap(VerificationDTO.class, Verification.class)
                .addMappings(m -> m.skip(Verification::setId));

        mapper.createTypeMap(FamilyMemberDTO.class, FamilyMember.class)
                .addMappings(m -> {
                    m.skip(FamilyMember::setId);
                    m.skip(FamilyMember::setApplication);
                });
    }

}
