from database.schema import connect, init_database


def install_audit_triggers():
    init_database()

    conn = connect()
    cur = conn.cursor()

    cur.executescript("""
    CREATE TRIGGER IF NOT EXISTS audit_stock_movements_insert
    AFTER INSERT ON stock_movements
    BEGIN
        INSERT INTO audit_logs (
            user_id, action, table_name, record_id, old_value, new_value
        )
        VALUES (
            NEW.user_id,
            'INSERT',
            'stock_movements',
            NEW.movement_id,
            NULL,
            'movement_type=' || COALESCE(NEW.movement_type, '') ||
            '; item_id=' || COALESCE(NEW.item_id, '') ||
            '; batch_id=' || COALESCE(NEW.batch_id, '') ||
            '; from_location_id=' || COALESCE(NEW.from_location_id, '') ||
            '; to_location_id=' || COALESCE(NEW.to_location_id, '') ||
            '; quantity=' || COALESCE(NEW.quantity, '') ||
            '; reference_no=' || COALESCE(NEW.reference_no, '') ||
            '; reason=' || COALESCE(NEW.reason, '')
        );
    END;


    CREATE TRIGGER IF NOT EXISTS audit_stock_movements_delete
    AFTER DELETE ON stock_movements
    BEGIN
        INSERT INTO audit_logs (
            user_id, action, table_name, record_id, old_value, new_value
        )
        VALUES (
            OLD.user_id,
            'DELETE',
            'stock_movements',
            OLD.movement_id,
            'movement_type=' || COALESCE(OLD.movement_type, '') ||
            '; item_id=' || COALESCE(OLD.item_id, '') ||
            '; batch_id=' || COALESCE(OLD.batch_id, '') ||
            '; quantity=' || COALESCE(OLD.quantity, '') ||
            '; reference_no=' || COALESCE(OLD.reference_no, '') ||
            '; reason=' || COALESCE(OLD.reason, ''),
            NULL
        );
    END;


    CREATE TRIGGER IF NOT EXISTS audit_items_insert
    AFTER INSERT ON items
    BEGIN
        INSERT INTO audit_logs (
            user_id, action, table_name, record_id, old_value, new_value
        )
        VALUES (
            NULL,
            'INSERT',
            'items',
            NEW.item_id,
            NULL,
            'item_code=' || COALESCE(NEW.item_code, '') ||
            '; generic_name=' || COALESCE(NEW.generic_name, '') ||
            '; brand_name=' || COALESCE(NEW.brand_name, '') ||
            '; dosage_form=' || COALESCE(NEW.dosage_form, '')
        );
    END;


    CREATE TRIGGER IF NOT EXISTS audit_batches_insert
    AFTER INSERT ON batches
    BEGIN
        INSERT INTO audit_logs (
            user_id, action, table_name, record_id, old_value, new_value
        )
        VALUES (
            NULL,
            'INSERT',
            'batches',
            NEW.batch_id,
            NULL,
            'item_id=' || COALESCE(NEW.item_id, '') ||
            '; batch_number=' || COALESCE(NEW.batch_number, '') ||
            '; expiry_date=' || COALESCE(NEW.expiry_date, '') ||
            '; supplier_id=' || COALESCE(NEW.supplier_id, '') ||
            '; received_date=' || COALESCE(NEW.received_date, '')
        );
    END;


    CREATE TRIGGER IF NOT EXISTS audit_suppliers_insert
    AFTER INSERT ON suppliers
    BEGIN
        INSERT INTO audit_logs (
            user_id, action, table_name, record_id, old_value, new_value
        )
        VALUES (
            NULL,
            'INSERT',
            'suppliers',
            NEW.supplier_id,
            NULL,
            'supplier_name=' || COALESCE(NEW.supplier_name, '') ||
            '; phone=' || COALESCE(NEW.phone, '') ||
            '; notes=' || COALESCE(NEW.notes, '')
        );
    END;


    CREATE TRIGGER IF NOT EXISTS audit_purchase_orders_insert
    AFTER INSERT ON purchase_orders
    BEGIN
        INSERT INTO audit_logs (
            user_id, action, table_name, record_id, old_value, new_value
        )
        VALUES (
            NEW.requested_by,
            'INSERT',
            'purchase_orders',
            NEW.po_id,
            NULL,
            'po_number=' || COALESCE(NEW.po_number, '') ||
            '; supplier_id=' || COALESCE(NEW.supplier_id, '') ||
            '; status=' || COALESCE(NEW.status, '')
        );
    END;


    CREATE TRIGGER IF NOT EXISTS audit_purchase_orders_update
    AFTER UPDATE ON purchase_orders
    BEGIN
        INSERT INTO audit_logs (
            user_id, action, table_name, record_id, old_value, new_value
        )
        VALUES (
            NEW.approved_by,
            'UPDATE',
            'purchase_orders',
            NEW.po_id,
            'status=' || COALESCE(OLD.status, ''),
            'status=' || COALESCE(NEW.status, '')
        );
    END;
    """)

    conn.commit()
    conn.close()


def show_audit_count():
    conn = connect()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM audit_logs;")
    count = cur.fetchone()[0]
    conn.close()
    print("Audit log rows:", count)


if __name__ == "__main__":
    install_audit_triggers()
    print("Audit triggers installed successfully.")
    show_audit_count()
