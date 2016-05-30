# Checking Modules
import customer_contract
import supplier_contract
import ttc_contract
import container_group
import module_group
import ttc_parts
import customer_contract_details
import supplier_parts
import customer_parts
import inner_packing_bom
import build_out

# function to select proper master check
def master_check(master, master_files, path):
    switcher = {
        'Build-out Master': build_out.build_out, # Start from row 10
        'Customer Contract Parts Master': customer_contract_details.customer_contract_details,
        'Supplier parts master': supplier_parts.supplier_parts,
        'Customer Parts Master': customer_parts.customer_parts,
        'Parts master': ttc_parts.ttc_parts,
        'Inner Packing BOM': inner_packing_bom.inner_packing_bom,
        'Module Group master': module_group.module_group,
        'Customer Contract master': customer_contract.customer_contract, # Start from row 11
        'Supplier Contract master': supplier_contract.supplier_contract,
        'TTC-Contract Master': ttc_contract.ttc_contract,
        'Container Group Master': container_group.container_group
    }
    # Get the function from switcher dictionary
    func = switcher.get(master, lambda: "nothing")
    # Execute the function
    return func(master_files, path)
